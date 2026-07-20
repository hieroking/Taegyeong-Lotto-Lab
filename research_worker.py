from __future__ import annotations
import hashlib, heapq, json, math, sys, time, traceback
from collections import Counter
from dataclasses import dataclass, asdict
from itertools import combinations
from pathlib import Path

PRICE_PER_LINE=1000
PAYOUT={6:1_500_000_000,"5b":50_000_000,5:1_300_000,4:5_000,3:1_000}

@dataclass(frozen=True)
class LottoRow:
    round_no:int
    numbers:tuple[int,int,int,int,int,int]
    bonus:int|None=None

@dataclass
class RoundResult:
    round_no:int
    max_match:int
    candidate_hits:int
    candidate_all6:int
    first:int
    second:int
    third:int
    fourth:int
    fifth:int
    winnings:int
    cost:int

@dataclass
class EngineSummary:
    engine:str
    rounds:int
    roi:float
    first:int
    second:int
    third:int
    fourth:int
    fifth:int
    avg_max_match:float
    max4_plus_rounds:int
    max5_plus_rounds:int
    exact6_rounds:int
    candidate_all6_rounds:int
    avg_candidate_hits:float
    reproducibility_hash:str

def emit(kind,**payload):
    print(json.dumps({"type":kind,**payload},ensure_ascii=False),flush=True)

def zscores(counter):
    vals=[counter.get(n,0) for n in range(1,46)]
    mean=sum(vals)/45
    sd=math.sqrt(sum((x-mean)**2 for x in vals)/45) or 1
    return {n:(counter.get(n,0)-mean)/sd for n in range(1,46)}

def gap_scores(history):
    last={n:-1 for n in range(1,46)}
    for i,draw in enumerate(history):
        for n in draw.numbers:
            last[n]=i
    gaps={n:len(history)-1-last[n] if last[n]>=0 else len(history) for n in range(1,46)}
    mean=sum(gaps.values())/45
    sd=math.sqrt(sum((g-mean)**2 for g in gaps.values())/45) or 1
    return {n:(gaps[n]-mean)/sd for n in range(1,46)}

def number_scores(history,engine):
    weights={
        "baseline":(.45,.30,.15,.10),
        "survivor":(.30,.30,.20,.20),
        "consensus":(.25,.35,.30,.10),
    }[engine]
    allc=Counter(n for draw in history for n in draw.numbers)
    c240=Counter(n for draw in history[-240:] for n in draw.numbers)
    c100=Counter(n for draw in history[-100:] for n in draw.numbers)
    za,z240,z100,g=zscores(allc),zscores(c240),zscores(c100),gap_scores(history)
    return {
        n:weights[0]*za[n]+weights[1]*z240[n]+weights[2]*z100[n]+weights[3]*g[n]
        for n in range(1,46)
    }

def pair_scores(history):
    counter=Counter()
    for draw in history[-240:]:
        counter.update(combinations(draw.numbers,2))
    vals=list(counter.values()) or [0]
    mean=sum(vals)/len(vals)
    sd=math.sqrt(sum((v-mean)**2 for v in vals)/len(vals)) or 1
    return {pair:(value-mean)/sd for pair,value in counter.items()}

def structure_score(combo):
    total=sum(combo)
    odd=sum(n%2 for n in combo)
    low=sum(n<=22 for n in combo)
    consecutive=sum(1 for a,b in zip(combo,combo[1:]) if b==a+1)
    same_end=6-len({n%10 for n in combo})
    return (
        -abs(total-138)/25
        -abs(odd-3)*.55
        -abs(low-3)*.45
        -max(0,consecutive-2)*1.2
        -max(0,same_end-2)*.7
    )

def generate_combinations(history,engine,count=100,candidate_size=18,stage=None):
    if stage:
        stage(0,100,"번호 점수 계산")
    scores=number_scores(history,engine)
    ranked=sorted(range(1,46),key=lambda n:(-scores[n],n))
    candidates=ranked[:candidate_size]
    pairs=pair_scores(history)
    total_combos=math.comb(len(candidates),6)
    keep=max(1200,min(5000,count*40))
    heap=[]

    for idx,combo in enumerate(combinations(candidates,6),1):
        number_score=sum(scores[n] for n in combo)
        pair_score=sum(pairs.get(tuple(sorted(p)),0) for p in combinations(combo,2))/15
        structural=structure_score(combo)
        if engine=="baseline":
            total=number_score+.35*pair_score+.8*structural
        elif engine=="survivor":
            total=.85*number_score+.25*pair_score+1.05*structural
        else:
            total=.75*number_score+.55*pair_score+.95*structural

        item=(total,tuple(-n for n in combo),combo)
        if len(heap)<keep:
            heapq.heappush(heap,item)
        elif item>heap[0]:
            heapq.heapreplace(heap,item)

        if idx%100==0:
            if stage:
                stage(idx,total_combos,f"조합 평가 {idx:,}/{total_combos:,}")
            time.sleep(.001)

    pool=[(score,combo) for score,_,combo in heap]
    pool.sort(key=lambda x:(-x[0],x[1]))
    if stage:
        stage(95,100,"추천 조합 다양성 정리")

    selected=[]
    for _,combo in pool:
        current=set(combo)
        if all(len(current.intersection(previous))<=4 for previous in selected[-30:]):
            selected.append(combo)
            if len(selected)>=count:
                break
    if len(selected)<count:
        used=set(selected)
        for _,combo in pool:
            if combo not in used:
                selected.append(combo)
                used.add(combo)
                if len(selected)>=count:
                    break
    if stage:
        stage(100,100,"조합 생성 완료")
    return candidates,selected

def evaluate(draw,candidates,combos):
    actual=set(draw.numbers)
    candidate_hits=len(actual.intersection(candidates))
    values=dict(first=0,second=0,third=0,fourth=0,fifth=0)
    max_match=0
    winnings=0
    for combo in combos:
        match=len(actual.intersection(combo))
        max_match=max(max_match,match)
        if match==6:
            values["first"]+=1; winnings+=PAYOUT[6]
        elif match==5 and draw.bonus in combo:
            values["second"]+=1; winnings+=PAYOUT["5b"]
        elif match==5:
            values["third"]+=1; winnings+=PAYOUT[5]
        elif match==4:
            values["fourth"]+=1; winnings+=PAYOUT[4]
        elif match==3:
            values["fifth"]+=1; winnings+=PAYOUT[3]
    return RoundResult(
        draw.round_no,max_match,candidate_hits,int(candidate_hits==6),
        **values,winnings=winnings,cost=len(combos)*PRICE_PER_LINE
    )

def run_engine(rows,engine,start,end,count,candidate_size,progress):
    by_round={row.round_no:row for row in rows}
    rounds=[round_no for round_no in range(start,end+1) if round_no in by_round]
    if not rounds:
        raise ValueError("검증 구간에 회차가 없습니다.")
    details=[]
    digest=hashlib.sha256()

    for i,round_no in enumerate(rounds,1):
        history=[row for row in rows if row.round_no<round_no]
        progress(i-1,len(rounds),engine,round_no,"회차 준비")
        def stage(done,total,message):
            progress(i-1+(done/max(total,1))*.90,len(rounds),engine,round_no,message)
        candidates,combos=generate_combinations(history,engine,count,candidate_size,stage)
        progress(i-.05,len(rounds),engine,round_no,"당첨 결과 비교")
        result=evaluate(by_round[round_no],candidates,combos)
        details.append(result)
        digest.update(f"{engine}|{round_no}|{candidates}|{combos}|{result.winnings}".encode())
        progress(i,len(rounds),engine,round_no,"회차 완료")

    cost=sum(x.cost for x in details)
    winnings=sum(x.winnings for x in details)
    summary=EngineSummary(
        engine,len(details),(winnings-cost)/cost*100 if cost else 0,
        sum(x.first for x in details),sum(x.second for x in details),
        sum(x.third for x in details),sum(x.fourth for x in details),
        sum(x.fifth for x in details),
        sum(x.max_match for x in details)/len(details),
        sum(x.max_match>=4 for x in details),
        sum(x.max_match>=5 for x in details),
        sum(x.max_match>=6 for x in details),
        sum(x.candidate_all6 for x in details),
        sum(x.candidate_hits for x in details)/len(details),
        digest.hexdigest()
    )
    return summary,details

def main():
    if len(sys.argv)!=2:
        raise SystemExit("input.json 경로가 필요합니다.")
    cfg=json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    rows=[
        LottoRow(
            int(x["round_no"]),
            tuple(int(n) for n in x["numbers"]),
            int(x["bonus"]) if x.get("bonus") is not None else None
        )
        for x in cfg["rows"]
    ]
    start=int(cfg["start"])
    end=int(cfg["end"])
    count=int(cfg["count"])
    candidate_size=int(cfg["candidate_size"])
    output_path=Path(cfg["output_path"])

    summaries=[]
    details={}
    engines=["baseline","survivor","consensus"]
    total_engines=len(engines)
    emit("progress",percent=2,message="자동연구 별도 프로세스 정상 시작")

    for engine_index,engine in enumerate(engines):
        emit("progress",percent=max(2,int(engine_index/total_engines*100)),message=f"{engine} 엔진 준비 중")
        def callback(done,total,current_engine,round_no,stage="검증 중"):
            fraction=max(0,min(1,float(done)/max(float(total),1)))
            percent=max(2,min(99,int(((engine_index+fraction)/total_engines)*100)))
            emit("progress",percent=percent,message=f"{current_engine} · {round_no}회 · {stage}")
        summary,round_details=run_engine(
            rows,engine,start,end,count,candidate_size,callback
        )
        summaries.append(summary)
        details[engine]=round_details
        emit("log",message=f"{engine} 엔진 완료 · ROI {summary.roi:.2f}%")

    baseline=summaries[0]
    decisions=[]
    for challenger in summaries[1:]:
        wins=sum([
            challenger.roi>baseline.roi,
            challenger.max4_plus_rounds>baseline.max4_plus_rounds,
            challenger.avg_max_match>baseline.avg_max_match,
            challenger.candidate_all6_rounds>baseline.candidate_all6_rounds
        ])
        high=(
            challenger.first+challenger.second+challenger.third
            >= baseline.first+baseline.second+baseline.third
        )
        status="채택 후보" if wins>=3 and high else "보류/폐기"
        decisions.append((challenger.engine,wins,high,status))

    output_path.write_text(json.dumps({
        "summaries":[asdict(x) for x in summaries],
        "details":{key:[asdict(x) for x in value] for key,value in details.items()},
        "decisions":decisions
    },ensure_ascii=False),encoding="utf-8")
    emit("progress",percent=100,message="자동연구 완료")
    return 0

if __name__=="__main__":
    try:
        raise SystemExit(main())
    except Exception:
        emit("error",message=traceback.format_exc())
        raise

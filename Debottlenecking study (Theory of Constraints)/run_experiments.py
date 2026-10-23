"""Run the rope-size sweep and all scenarios; write results.json."""
import json, time
import simulate as S

SEEDS = [2026 + k for k in range(8)]
SWEEP_SEEDS = SEEDS[:5]
ROPES = [150, 250, 400, 600, 900, 1500]
t0 = time.time()

sweep = []
for K in ROPES:
    runs = [S.run(["DBR"], sd, rope_units=K) for sd in SWEEP_SEEDS]
    sm = S.summarise(runs)
    sweep.append(dict(rope=K, good_wk=sm["good_wk"], wip=sm["wip"], flow_hr=sm["flow_hr"], anod_util=sm["util"][S.ANOD]))
    print("rope", K, round(sm["good_wk"]["mean"]), round(sm["wip"]["mean"]), round(sm["flow_hr"]["mean"], 1), flush=True)
best = max(s["good_wk"]["mean"] for s in sweep)
chosen = next(s["rope"] for s in sweep if s["good_wk"]["mean"] >= 0.998 * best)
print("chosen rope", chosen, "elapsed", round(time.time() - t0), flush=True)

# second sweep: once the constraint is faster, the protective buffer must be re-sized
ROPES2 = [400, 600, 900, 1500]
sweep2 = []
for K in ROPES2:
    runs = [S.run(["DBR", "SMED"], sd, rope_units=K) for sd in SWEEP_SEEDS]
    sm = S.summarise(runs)
    sweep2.append(dict(rope=K, good_wk=sm["good_wk"], wip=sm["wip"], flow_hr=sm["flow_hr"], anod_util=sm["util"][S.ANOD]))
    print("rope2", K, round(sm["good_wk"]["mean"]), round(sm["wip"]["mean"]), flush=True)
best2 = max(s_["good_wk"]["mean"] for s_ in sweep2)
chosen2 = next(s_["rope"] for s_ in sweep2 if s_["good_wk"]["mean"] >= 0.998 * best2)
print("chosen rope2", chosen2, flush=True)

scen = []
for sid, name, flags in S.SCENARIOS:
    rope = chosen if flags in ([], ["DBR"]) else chosen2
    runs = [S.run(flags, sd, rope_units=rope) for sd in SEEDS]
    sm = S.summarise(runs)
    stages, policy = S.apply_flags(flags)
    scen.append(dict(id=sid, name=name, flags=flags, policy=policy, rope=(rope if policy == "DBR" else None), summary=sm))
    print(sid, name, round(sm["good_wk"]["mean"]), "+/-", round(sm["good_wk"]["ci95"]), "WIP", round(sm["wip"]["mean"]),
          "flow h", round(sm["flow_hr"]["mean"], 1), "anod util", round(sm["util"][S.ANOD], 2), flush=True)

json.dump(dict(seeds=SEEDS, sweep_seeds=SWEEP_SEEDS, rope_chosen=chosen, rope_resized=chosen2, sweep=sweep, sweep2=sweep2, scenarios=scen,
               params=dict(lot=S.LOT, rack=S.RACK, t_week=S.T_WEEK, demand=S.DEMAND, cv=S.CV, weeks=8, warm_weeks=2,
                           stages=S.BASE_STAGES, fix=S.FIX_PARAMS)),
          open("/home/claude/results.json", "w"), indent=1)
print("done", round(time.time() - t0), "s")

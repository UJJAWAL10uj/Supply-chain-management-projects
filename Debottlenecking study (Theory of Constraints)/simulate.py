"""
Theory-of-Constraints debottlenecking study: SimPy model of a multi-stage aluminium-enclosure line.

Line (one production module, 10,000 good units/week target):
  Blank saw -> CNC OP10 rough -> CNC OP20 finish -> Deburr/polish -> Anodising (batch racks)
  -> Inspection -> Packing

DATA IS SIMULATED / ILLUSTRATIVE (not a real plant). Stage parameters are plausible benchmark-style
values chosen so the line behaves like a typical enclosure line; replace them with time-study data.

Model features: per-machine breakdowns (operating-time MTBF / MTTR), cycle-time variability,
stage yield loss, batch anodising with 48-unit racks, push release at demand rate or
drum-buffer-rope (DBR) release gated by a buffer ahead of the constraint.
"""
import json
import math
import random
import statistics
import sys
import time

import simpy

LOT = 6                      # units per transfer lot (tray)
RACK = 48                    # units per anodising rack
T_WEEK = 96 * 60.0           # planned production minutes per week (2 shifts x 8 h x 6 days)
DEMAND = 10000               # good units per week (customer demand)
CV = 0.15                    # coefficient of variation of lot processing time

# name, kind, machines/lines, ideal cycle min per unit (batch: per rack in 'batch_min'), availability,
# performance, yield, MTTR (min)
BASE_STAGES = [
    dict(name="Blank saw", kind="unit", n=1, cycle=0.30, A=0.92, P=0.95, Y=0.995, mttr=20),
    dict(name="CNC OP10 rough", kind="unit", n=7, cycle=2.60, A=0.90, P=0.92, Y=0.990, mttr=25),
    dict(name="CNC OP20 finish", kind="unit", n=9, cycle=3.40, A=0.90, P=0.92, Y=0.975, mttr=25),
    dict(name="Deburr / polish", kind="unit", n=5, cycle=1.80, A=0.95, P=0.90, Y=0.990, mttr=15),
    dict(name="Anodising (rack batch)", kind="batch", n=2, cycle=None, batch_min=52.0, A=0.93, P=1.00, Y=0.905, mttr=40),
    dict(name="Inspection", kind="unit", n=2, cycle=0.85, A=0.96, P=0.95, Y=0.999, mttr=10),
    dict(name="Packing", kind="unit", n=2, cycle=0.50, A=0.97, P=0.95, Y=0.999, mttr=10),
]
ANOD = 4  # index of the anodising stage

# Improvement levers on the anodising constraint
FIX_PARAMS = dict(yield_new=0.935, smed_batch_min=46.0, sevenday_mult=7.0 / 6.0, line3_lines=3, tpm_A=0.96)

FLAGS = ["DBR", "YIELD", "SMED", "SEVEN", "LINE3", "TPM"]
SCENARIOS = [
    ("C0", "Base case (push release)", []),
    ("C1", "DBR release only", ["DBR"]),
    ("C2", "Anodising yield 90.5% to 93.5%", ["YIELD"]),
    ("C3", "SMED: rack batch 52 to 46 min", ["SMED"]),
    ("C4", "Anodising 7 days a week", ["SEVEN"]),
    ("C5", "Add third anodising line", ["LINE3"]),
    ("C6", "TPM: anodising availability 96%", ["TPM"]),
    ("C7", "Low-cost package: DBR + SMED", ["DBR", "SMED"]),
    ("C8", "DBR + SMED + yield at constraint", ["DBR", "SMED", "YIELD"]),
    ("C9", "DBR + SMED + TPM", ["DBR", "SMED", "TPM"]),
    ("C10", "Capex route: DBR + third line", ["DBR", "LINE3"]),
    ("C11", "All fixes", ["DBR", "YIELD", "SMED", "SEVEN", "LINE3", "TPM"]),
]


def apply_flags(flags):
    """Return (stages, policy) for a scenario's flag list."""
    stages = [dict(s) for s in BASE_STAGES]
    a = stages[ANOD]
    if "YIELD" in flags:
        a["Y"] = FIX_PARAMS["yield_new"]
    if "SMED" in flags:
        a["batch_min"] = FIX_PARAMS["smed_batch_min"]
    if "LINE3" in flags:
        a["n"] = FIX_PARAMS["line3_lines"]
    if "TPM" in flags:
        a["A"] = FIX_PARAMS["tpm_A"]
    a["time_mult"] = FIX_PARAMS["sevenday_mult"] if "SEVEN" in flags else 1.0
    policy = "DBR" if "DBR" in flags else "PUSH"
    return stages, policy


def total_yield(stages):
    y = 1.0
    for s in stages:
        y *= s["Y"]
    return y


class Tracker:
    def __init__(self, env, n_stages, warm):
        self.env, self.warm = env, warm
        self.units = [0] * n_stages
        self.integral = [0.0] * n_stages
        self.last = [0.0] * n_stages
        self.busy = [0.0] * n_stages
        self.down = [0.0] * n_stages
        self.busy_snap = None
        self.down_snap = None
        self.good = 0
        self.good_snap = 0
        self.pool_level = 0

    def change(self, i, delta):
        now = self.env.now
        t0 = max(self.last[i], self.warm)
        if now > t0:
            self.integral[i] += self.units[i] * (now - t0)
        self.last[i] = now
        self.units[i] += delta

    def snapshot(self):
        self.busy_snap = list(self.busy)
        self.down_snap = list(self.down)
        self.good_snap = self.good

    def ahead(self):
        return self.units[0] + self.units[1] + self.units[2] + self.units[3]


def lot_time(rng, mean):
    sigma = math.sqrt(math.log(1 + CV * CV))
    return mean * rng.lognormvariate(-sigma * sigma / 2, sigma)


def binom(rng, n, p):
    return sum(1 for _ in range(n) if rng.random() < p)


def run(flags, seed, weeks=8, warm_weeks=2, rope_units=None, force_policy=None, release_mult=1.0):
    stages, policy = apply_flags(flags)
    if force_policy:
        policy = force_policy
    rng = random.Random(seed)
    env = simpy.Environment()
    warm = warm_weeks * T_WEEK
    horizon = weeks * T_WEEK
    trk = Tracker(env, len(stages), warm)
    queues = [simpy.Store(env) for _ in stages]
    pool = simpy.Container(env, init=0)  # units waiting at anodising for a rack

    def send(i, lot):
        trk.change(i, lot)
        if i == ANOD:
            return pool.put(lot)
        return queues[i].put(lot)

    def breakdown_time(rng_, remaining, state, i, mtbf, mttr):
        """Generator: processes `remaining` minutes of work with operating-time failures."""
        while remaining > 1e-9:
            if state["ttf"] < remaining:
                yield env.timeout(state["ttf"])
                trk.busy[i] += state["ttf"]
                remaining -= state["ttf"]
                rep = rng_.expovariate(1.0 / mttr)
                yield env.timeout(rep)
                trk.down[i] += rep
                state["ttf"] = rng_.expovariate(1.0 / mtbf)
            else:
                yield env.timeout(remaining)
                trk.busy[i] += remaining
                state["ttf"] -= remaining
                remaining = 0

    def unit_worker(i, s, wrng):
        mtbf = s["A"] / (1 - s["A"]) * s["mttr"]
        state = dict(ttf=wrng.expovariate(1.0 / mtbf))
        out_i = i + 1
        while True:
            units = yield queues[i].get()
            mean = units * s["cycle"] / s["P"]
            yield from breakdown_time(wrng, lot_time(wrng, mean), state, i, mtbf, s["mttr"])
            good = binom(wrng, units, s["Y"])
            trk.change(i, -units)
            if out_i < len(stages):
                if good:
                    yield send(out_i, good)
            else:
                trk.good += good

    def anod_worker(i, s, wrng):
        mtbf = s["A"] / (1 - s["A"]) * s["mttr"]
        state = dict(ttf=wrng.expovariate(1.0 / mtbf))
        while True:
            yield pool.get(RACK)
            mean = s["batch_min"] / s.get("time_mult", 1.0) / s["P"]
            yield from breakdown_time(wrng, lot_time(wrng, mean), state, i, mtbf, s["mttr"])
            good = binom(wrng, RACK, s["Y"])
            trk.change(i, -RACK)
            while good > 0:
                chunk = min(LOT, good)
                good -= chunk
                yield send(i + 1, chunk)

    for i, s in enumerate(stages):
        for k in range(s["n"]):
            wrng = random.Random(rng.random())
            env.process(anod_worker(i, s, wrng) if s["kind"] == "batch" else unit_worker(i, s, wrng))

    # release rate = demand divided by the planner's expected yield
    starts_per_min = release_mult * DEMAND / total_yield(stages) / T_WEEK
    interval = LOT / starts_per_min
    gate = rope_units if rope_units is not None else 600

    def source():
        while True:
            yield env.timeout(interval)
            if policy == "DBR":
                while trk.ahead() + pool.level >= gate:
                    yield env.timeout(1.0)
            yield send(0, LOT)

    def sampler():
        yield env.timeout(warm)
        trk.snapshot()

    env.process(source())
    env.process(sampler())
    env.run(until=horizon)
    for i in range(len(stages)):
        trk.change(i, 0)  # flush integrals

    window = horizon - warm
    wk = window / T_WEEK
    good_wk = (trk.good - trk.good_snap) / wk
    wip = [trk.integral[i] / window for i in range(len(stages))]
    util = [(trk.busy[i] - trk.busy_snap[i]) / (stages[i]["n"] * window) for i in range(len(stages))]
    wip_total = sum(wip)
    flow_hr = wip_total / (good_wk / (T_WEEK / 60.0)) if good_wk > 0 else float("nan")
    return dict(good_wk=good_wk, wip=wip_total, wip_by_stage=wip, util=util, flow_hr=flow_hr)


def summarise(runs):
    def ms(vals):
        m = statistics.mean(vals)
        sd = statistics.stdev(vals) if len(vals) > 1 else 0.0
        ci = 2.365 * sd / math.sqrt(len(vals)) if len(vals) > 1 else 0.0  # t(0.975, 7)
        return m, sd, ci
    out = {}
    for key in ("good_wk", "wip", "flow_hr"):
        m, sd, ci = ms([r[key] for r in runs])
        out[key] = dict(mean=m, sd=sd, ci95=ci)
    out["util"] = [statistics.mean(r["util"][i] for r in runs) for i in range(len(runs[0]["util"]))]
    out["wip_by_stage"] = [statistics.mean(r["wip_by_stage"][i] for r in runs) for i in range(len(runs[0]["util"]))]
    return out


if __name__ == "__main__":
    reps = int(sys.argv[1]) if len(sys.argv) > 1 else 8
    seeds = [2026 + k for k in range(reps)]
    t0 = time.time()
    r = run([], seeds[0])
    print("single base run %.1fs" % (time.time() - t0), round(r["good_wk"]), round(r["wip"]), [round(u, 2) for u in r["util"]])

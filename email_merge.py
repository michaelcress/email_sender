import json, time, random, subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

MAX_CONCURRENCY = 3            # threads in parallel (tune carefully)
RATE_PER_MINUTE = 60           # hard cap across all threads
SPACING = 60.0 / max(1, RATE_PER_MINUTE)

def send_one(rec):
    # light jitter to avoid thundering herd
    time.sleep(random.uniform(0, SPACING))

    # build your CLI invocation (example args)
    cmd = [
        "email-sender",                      # your email CLI
        "--to", rec["email"],
        "--subject", rec["subject"],
        "--body", rec["body"],               # or --body-file path
    ]

    backoff = 1.0
    for attempt in range(1, 6):  # up to 5 tries
        t0 = time.time()
        p = subprocess.run(cmd, capture_output=True, text=True)
        dt = time.time() - t0

        if p.returncode == 0:
            return {"email": rec["email"], "ok": True, "ms": int(dt*1000)}
        # transient? back off and retry
        time.sleep(backoff + random.uniform(0, 0.200))
        backoff = min(backoff * 2, 16.0)

    return {
        "email": rec["email"],
        "ok": False,
        "error": (p.stderr or p.stdout).strip()[:2000]
    }

def run_mail_merge(records):
    results = []
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as pool:
        futures = [pool.submit(send_one, r) for r in records]
        for fut in as_completed(futures):
            res = fut.result()
            print(json.dumps(res, ensure_ascii=False))
            results.append(res)
    return results

# Example usage:
# records = [{"email":"a@x.com","subject":"Hi","body":"..."},
#            {"email":"b@y.com","subject":"Hi","body":"..."}]
# run_mail_merge(records)

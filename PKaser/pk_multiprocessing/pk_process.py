from threading import Thread


def runWithThreads(*fns):
    proc = []
    for fn in fns:
        p = Thread(target=fn)
        print(p)
        p.start()
        proc.append(p)
    for p in proc:
        print(p)
        p.join()

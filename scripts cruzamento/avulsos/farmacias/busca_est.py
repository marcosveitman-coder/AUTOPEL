import pickle, glob, os, time
import busca_assoc as b
from multiprocessing import Pool
def init(r): b.ROOTS = r
if __name__ == "__main__":
    t = time.time(); roots = set(pickle.load(open("busca_emp.pkl", "rb")))
    out = []
    with Pool(4, initializer=init, initargs=(roots,)) as pool:
        for o in pool.imap_unordered(b.est, sorted(glob.glob(os.path.join(b.DIR, "Estabelecimentos*.zip")))): out += o
    pickle.dump(out, open("busca_est.pkl", "wb")); print("matrizes:", len(out), f"{time.time()-t:.0f}s")

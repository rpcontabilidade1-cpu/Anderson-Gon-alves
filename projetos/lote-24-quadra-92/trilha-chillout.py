# Trilha original no estilo chill out balear, 90,5 s, sincronizada com o vídeo.
import numpy as np, wave, sys

SR = 44100
DUR = 90.5
N = int(SR * DUR)
BPM = 98.0
BEAT = 60.0 / BPM
BAR = 4 * BEAT
rng = np.random.default_rng(7)
L = np.zeros(N); R = np.zeros(N)

def t_(d): return np.arange(int(SR * d)) / SR
def env(n, a, r, s=1.0):
    e = np.ones(n) * s
    na, nr = int(SR * a), int(SR * r)
    na = min(na, n); e[:na] = np.linspace(0, s, na)
    if nr and nr < n: e[-nr:] *= np.linspace(1, 0, nr)
    return e
def put(sig, t0, pan=0.0, gain=1.0):
    i = int(SR * t0)
    if i >= N: return
    sig = sig[:N - i] * gain
    L[i:i + len(sig)] += sig * np.sqrt(0.5 * (1 - pan))
    R[i:i + len(sig)] += sig * np.sqrt(0.5 * (1 + pan))
def nota(m): return 440.0 * 2 ** ((m - 69) / 12)
def lp(x, a):                                  # passa-baixa de um polo
    y = np.empty_like(x); acc = 0.0
    for k in range(len(x)): acc += a * (x[k] - acc); y[k] = acc
    return y
def lp_fast(x, a):
    k = int(np.ceil(np.log(1e-4) / np.log(1 - a)))
    h = a * (1 - a) ** np.arange(k)
    n = len(x) + k - 1; nf = 1 << (n - 1).bit_length()
    return np.fft.irfft(np.fft.rfft(x, nf) * np.fft.rfft(h, nf), nf)[:len(x)]

# progressão: Am9 | Fmaj9 | Cmaj9 | G6/B   (2 compassos cada)
ACORDES = [[57, 60, 64, 67, 71], [53, 57, 60, 64, 67], [48, 55, 59, 62, 64], [47, 55, 59, 62, 64]]
BAIXO = [45, 41, 36, 43]

# seções (s): intro 0-7 | groove 7-66 | respiro 66-79 | final 79-86 | coda 86-90.5
def secao(t):
    if t < 7.0: return "intro"
    if t < 66.0: return "groove"
    if t < 79.0: return "respiro"
    if t < 86.0: return "final"
    return "coda"

# --- pad (dente de serra desafinado + filtro) ---
def pad(midis, d):
    tt = t_(d); s = np.zeros(len(tt))
    for m in midis:
        for det in (-0.08, 0.0, 0.07):
            f = nota(m + det)
            s += 2 * ((tt * f + rng.random()) % 1.0) - 1
    s /= (3 * len(midis))
    s = lp_fast(s, 0.035)
    lfo = 1 + 0.15 * np.sin(2 * np.pi * 0.2 * tt)
    return s * env(len(tt), 1.2, 1.4) * lfo

# --- piano elétrico (FM simples) ---
def ep(m, d=1.6, vel=1.0):
    tt = t_(d); f = nota(m)
    mod = np.sin(2 * np.pi * f * tt) * 1.6 * np.exp(-tt * 3)
    s = np.sin(2 * np.pi * f * tt + mod) * np.exp(-tt * 2.2)
    s += 0.25 * np.sin(2 * np.pi * 2 * f * tt) * np.exp(-tt * 5)
    return s * env(len(tt), 0.004, 0.3) * vel

# --- baixo ---
def baixo(m, d):
    tt = t_(d); f = nota(m)
    s = np.sin(2 * np.pi * f * tt) + 0.3 * np.sin(2 * np.pi * 2 * f * tt)
    return s * env(len(tt), 0.01, 0.15)

# --- bateria ---
def kick():
    tt = t_(0.45); f = 50 + 70 * np.exp(-tt * 25)
    ph = 2 * np.pi * np.cumsum(f) / SR
    return np.sin(ph) * np.exp(-tt * 7)
def hat(d=0.06):
    n = rng.standard_normal(int(SR * d)); n = n - lp_fast(n, 0.5)
    return n * np.exp(-np.arange(len(n)) / SR * 70)
def shaker():
    n = rng.standard_normal(int(SR * 0.09)); n = n - lp_fast(n, 0.3)
    e = np.sin(np.linspace(0, np.pi, len(n))) ** 2
    return n * e
def clap():
    n = rng.standard_normal(int(SR * 0.25)); n = lp_fast(n - lp_fast(n, 0.08), 0.6)
    return n * np.exp(-np.arange(len(n)) / SR * 18)

# --- mar (ruído filtrado com ondulação lenta) ---
_r = rng.standard_normal(N); mar = (lp_fast(_r, 0.22) - lp_fast(_r, 0.035)) * 0.6
tt = np.arange(N) / SR
onda = 0.5 + 0.5 * np.sin(2 * np.pi * tt / 7.5) ** 2
mar_amp = np.interp(tt, [0, 6, 9, 62, 70, 80, 86, DUR], [0.9, 0.9, 0.25, 0.25, 0.7, 0.4, 0.6, 0.9])
L += mar * onda * mar_amp * 0.9
R += np.roll(mar, 1500) * onda * mar_amp * 0.9

# --- montagem por compasso ---
ARP = [0, 2, 4, 3, 1, 4, 2, 3]
nbar = int(DUR / BAR) + 1
for b in range(nbar):
    t0 = b * BAR
    if t0 >= DUR: break
    sec = secao(t0)
    c = ACORDES[(b // 2) % 4]; bx = BAIXO[(b // 2) % 4]
    if b % 2 == 0:
        put(pad(c, 2 * BAR + 1.0), t0, 0.0, 0.55 if sec != "groove" else 0.42)
    if sec in ("groove", "final"):
        put(baixo(bx, BEAT * 1.5), t0, 0, 0.22); put(baixo(bx, BEAT * 0.9), t0 + BEAT * 2.5, 0, 0.18)
        put(baixo(bx + 7, BEAT * 0.9), t0 + BEAT * 3.5, 0, 0.14)
        for k in range(4): put(kick(), t0 + k * BEAT, 0, 0.42)
        for k in range(4): put(hat(), t0 + k * BEAT + BEAT / 2, 0.35, 0.16)
        for k in range(8): put(shaker(), t0 + k * BEAT / 2 + BEAT / 4, -0.4, 0.06)
        if t0 > 18: put(clap(), t0 + BEAT, 0.1, 0.12); put(clap(), t0 + 3 * BEAT, 0.1, 0.12)
    if sec in ("groove", "respiro", "final") or (sec == "intro" and t0 > 2.4):
        for k, idx in enumerate(ARP):
            if sec == "intro" and k % 2: continue
            m = c[idx] + 12
            put(ep(m, 1.4, 0.8 if k % 2 == 0 else 0.55), t0 + k * BEAT / 2, (-0.5 if k % 2 else 0.5), 0.16)
    if sec == "coda" and b % 1 == 0:
        put(ep(57 + 12, 3.5, 1.0), t0, -0.3, 0.2); put(ep(64 + 12, 3.5, 0.8), t0 + BEAT, 0.3, 0.17)

# --- eco estéreo (dotted 1/8) e reverb simples ---
def eco(x, d, fb, mix):
    n = int(SR * d); y = x.copy()
    for k in range(1, 6): y[n * k:] += x[:-n * k] * (fb ** k) * mix if n * k < len(x) else 0
    return y
L2 = eco(L, BEAT * 0.75, 0.45, 0.35); R2 = eco(R, BEAT * 0.75 + 0.011, 0.45, 0.35)
def reverb(x):
    y = x.copy()
    for d, g in ((0.0297, .5), (0.0371, .45), (0.0411, .42), (0.0437, .4), (0.061, .3), (0.083, .25)):
        n = int(SR * d); z = np.zeros_like(x)
        for k in range(1, 8):
            if n * k < len(x): z[n * k:] += x[:-n * k] * g ** k
        y += lp_fast(z, 0.25) * 0.35
    return y
L2, R2 = reverb(L2), reverb(R2)

# --- master: fade, normalização, soft clip ---
fade = np.interp(tt, [0, 2.5, DUR - 4.5, DUR], [0, 1, 1, 0])
st = np.stack([L2 * fade, R2 * fade], 1)
st /= np.max(np.abs(st)) + 1e-9
st = np.tanh(st * 1.4) / np.tanh(1.4) * 0.89
pcm = (st * 32767).astype(np.int16)
with wave.open(sys.argv[1], "wb") as w:
    w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes(pcm.tobytes())
print("ok", len(pcm) / SR)

"""
generate_sounds.py
------------------
Run this ONCE before building the exe.
Produces 4 .wav files in the sounds/ subfolder.
PyInstaller will bundle them via the spec file.

Usage:  python generate_sounds.py
"""

import struct, math, wave, os
from pathlib import Path

OUT_DIR = Path(__file__).parent / "sounds"
OUT_DIR.mkdir(exist_ok=True)

SAMPLE_RATE = 44100

# ── helpers ────────────────────────────────────────────────────────────────────

def note_freq(name: str) -> float:
    """Return Hz for a note name like 'C4', 'E4', 'G4', 'A4', etc."""
    notes = ["C","C#","D","D#","E","F","F#","G","G#","A","A#","B"]
    name  = name.strip()
    octave = int(name[-1])
    pitch  = notes.index(name[:-1])
    # A4 = 440 Hz, MIDI note 69
    midi = (octave + 1) * 12 + pitch
    return 440.0 * (2 ** ((midi - 69) / 12))

def sine_wave(freq: float, duration: float, volume: float = 0.6,
              attack: float = 0.02, release: float = 0.08) -> list:
    """Generate a sine wave sample list with attack/release envelope."""
    n_samples = int(SAMPLE_RATE * duration)
    atk = int(SAMPLE_RATE * attack)
    rel = int(SAMPLE_RATE * release)
    samples = []
    for i in range(n_samples):
        t    = i / SAMPLE_RATE
        amp  = volume
        if i < atk:
            amp *= i / atk
        elif i > n_samples - rel:
            amp *= (n_samples - i) / rel
        samples.append(amp * math.sin(2 * math.pi * freq * t))
    return samples

def bell_wave(freq: float, duration: float, volume: float = 0.6) -> list:
    """Bell-like tone: fundamental + harmonics with exponential decay."""
    n_samples = int(SAMPLE_RATE * duration)
    harmonics = [(1.0, 1.0), (2.756, 0.45), (5.404, 0.22), (8.933, 0.10)]
    samples   = []
    decay     = 4.0   # decay rate
    for i in range(n_samples):
        t   = i / SAMPLE_RATE
        env = math.exp(-decay * t)
        s   = 0.0
        for ratio, amp in harmonics:
            s += amp * math.sin(2 * math.pi * freq * ratio * t)
        samples.append(volume * env * s / sum(a for _, a in harmonics))
    return samples

def silence(duration: float) -> list:
    return [0.0] * int(SAMPLE_RATE * duration)

def mix(a: list, b: list) -> list:
    """Mix two sample lists (pad shorter with zeros)."""
    length = max(len(a), len(b))
    result = []
    for i in range(length):
        va = a[i] if i < len(a) else 0.0
        vb = b[i] if i < len(b) else 0.0
        result.append(max(-1.0, min(1.0, va + vb)))
    return result

def save_wav(filename: str, samples: list):
    path = OUT_DIR / filename
    with wave.open(str(path), "w") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(SAMPLE_RATE)
        packed = struct.pack(f"<{len(samples)}h",
                             *[int(s * 32767) for s in samples])
        wf.writeframes(packed)
    print(f"  ✔  {path.name}  ({len(samples)/SAMPLE_RATE:.2f}s)")

# ── 1. Chime  (gentle ascending bell: C5 → E5 → G5) ──────────────────────────

def make_chime():
    c5 = bell_wave(note_freq("C5"), 1.0, 0.55)
    e5 = bell_wave(note_freq("E5"), 1.0, 0.55)
    g5 = bell_wave(note_freq("G5"), 1.2, 0.55)

    # stagger the notes with overlap
    gap = int(SAMPLE_RATE * 0.55)
    total = len(c5) + gap * 2 + len(g5)
    buf   = [0.0] * total

    for i, s in enumerate(c5):
        buf[i] += s
    for i, s in enumerate(e5):
        buf[i + gap] += s
    for i, s in enumerate(g5):
        buf[i + gap * 2] += s

    # normalise
    peak = max(abs(x) for x in buf) or 1
    buf  = [x / peak * 0.85 for x in buf]
    save_wav("chime.wav", buf)

# ── 2. Urgent  (rapid descending: G5 → E5 → C5, with edge) ───────────────────

def make_urgent():
    # short sharp tones with a slight buzz (add 2nd harmonic for edge)
    def sharp(freq, dur):
        n  = int(SAMPLE_RATE * dur)
        rel = int(SAMPLE_RATE * 0.04)
        s  = []
        for i in range(n):
            t   = i / SAMPLE_RATE
            env = 1.0 if i < n - rel else (n - i) / rel
            s.append(env * (0.7 * math.sin(2*math.pi*freq*t)
                          + 0.3 * math.sin(2*math.pi*freq*2*t)))
        return s

    g5 = sharp(note_freq("G5"), 0.18)
    e5 = sharp(note_freq("E5"), 0.18)
    c5 = sharp(note_freq("C5"), 0.50)
    gap = int(SAMPLE_RATE * 0.04)

    buf = g5 + [0.0]*gap + e5 + [0.0]*gap + c5
    # repeat once with short gap for urgency
    buf = buf + [0.0]*int(SAMPLE_RATE*0.25) + buf

    peak = max(abs(x) for x in buf) or 1
    buf  = [x / peak * 0.85 for x in buf]
    save_wav("urgent.wav", buf)

# ── 3. Doorbell  (ding-dong: G5 down to C5, classic two-note) ─────────────────

def make_doorbell():
    ding = bell_wave(note_freq("G5"), 1.3, 0.65)
    dong = bell_wave(note_freq("C5"), 1.5, 0.65)

    gap  = int(SAMPLE_RATE * 0.55)
    total = len(ding) + gap + len(dong)
    buf   = [0.0] * total

    for i, s in enumerate(ding):
        buf[i] += s
    for i, s in enumerate(dong):
        buf[i + gap + len(ding) - int(SAMPLE_RATE*0.3)] += s   # slight overlap

    peak = max(abs(x) for x in buf) or 1
    buf  = [x / peak * 0.85 for x in buf]
    save_wav("doorbell.wav", buf)

# ── 4. Fanfare  (4-note rising: C5 → E5 → G5 → C6) ──────────────────────────

def make_fanfare():
    def brass(freq, dur, vol=0.5):
        """Brass-ish: sine + odd harmonics."""
        n   = int(SAMPLE_RATE * dur)
        atk = int(SAMPLE_RATE * 0.03)
        rel = int(SAMPLE_RATE * 0.12)
        s   = []
        for i in range(n):
            t   = i / SAMPLE_RATE
            env = 1.0
            if i < atk:
                env = i / atk
            elif i > n - rel:
                env = (n - i) / rel
            s.append(vol * env * (
                0.6 * math.sin(2*math.pi*freq*t)
              + 0.25* math.sin(2*math.pi*freq*3*t)
              + 0.10* math.sin(2*math.pi*freq*5*t)
              + 0.05* math.sin(2*math.pi*freq*7*t)
            ))
        return s

    notes = [
        brass(note_freq("C5"), 0.30),
        brass(note_freq("E5"), 0.30),
        brass(note_freq("G5"), 0.30),
        brass(note_freq("C6"), 0.90),
    ]
    gap = int(SAMPLE_RATE * 0.02)
    buf = []
    for n in notes:
        buf.extend(n)
        buf.extend([0.0] * gap)

    peak = max(abs(x) for x in buf) or 1
    buf  = [x / peak * 0.85 for x in buf]
    save_wav("fanfare.wav", buf)

# ── main ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Generating built-in sounds…")
    make_chime()
    make_doorbell()
    make_fanfare()
    make_urgent()
    print(f"\nDone — files saved to:  {OUT_DIR}")

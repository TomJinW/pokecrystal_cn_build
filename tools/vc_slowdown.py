#!/usr/bin/env python3
"""
Append "SlowDown" entries to the Crystal CHS VC patch (pokecrystal11.patch).

The Chinese build runs in GBC double speed (8 MHz) all the time. On 3DS VC
every busy-wait loop then costs twice the emulated cycles. The official
Korean Gold/Silver VC patch (CGBAAUK0.959.patch) solves the same problem by
inserting `call DelayFrame` (halt until VBlank) into busy-wait loops.

This script does the same purely at the .patch level (no game source change),
using only Mode 1 entries (static ROM byte replacement at load time):

  * small trampolines are written into unused ROM0 space
  * the busy-wait instruction is replaced by `call <trampoline>`

Usage:  python3 tools/vc_slowdown.py [build_dir] [rom_base]
        (defaults: build pokecrystal11) -- run after `make crystal11_vc`.
Re-running is safe: the previous SlowDown block is replaced.
"""
import re, sys, os

build = sys.argv[1] if len(sys.argv) > 1 else 'build'
base  = sys.argv[2] if len(sys.argv) > 2 else 'pokecrystal11'
SYM, MAP, ROM, PATCH = (os.path.join(build, base + ext) for ext in ('.sym', '.map', '.gbc', '.patch'))

BEGIN = '; >>> VC SlowDown (tools/vc_slowdown.py) >>>'
END   = '; <<< VC SlowDown <<<'

def die(msg):
    sys.exit('vc_slowdown: ' + msg)

# ---- symbols -----------------------------------------------------------
syms = {}
for line in open(SYM, encoding='utf-8', errors='replace'):
    m = re.match(r'([0-9a-f]+):([0-9a-f]+) (\S+)', line)
    if m:
        syms.setdefault(m.group(3), (int(m.group(1), 16), int(m.group(2), 16)))

def sym(name):
    if name not in syms:
        die('symbol not found: ' + name)
    return syms[name]

def off(bank, addr):
    return addr if bank == 0 else bank * 0x4000 + addr - 0x4000

rom = open(ROM, 'rb').read()

def expect(o, data, what):
    if rom[o:o + len(data)] != data:
        die('%s: unexpected bytes at 0x%X: %s (expected %s)' % (
            what, o, rom[o:o + len(data)].hex(' '), data.hex(' ')))

le = lambda v: bytes([v & 0xFF, v >> 8])

_, DelayFrame        = sym('DelayFrame')
_, wTextDelayFrames  = sym('wTextDelayFrames')
_, w2DMenuFlags1     = sym('w2DMenuFlags1')
b_wait, a_wait       = sym('PrintLetterDelay.wait')
b_menu, a_menu       = sym('Do2DMenuRTCJoypad.loopRTC')
b_rel,  a_rel        = sym('ReloadTilesetAndPalettes')
_, SkipMusic         = sym('SkipMusic')

# ---- trampolines -------------------------------------------------------
# T1: call DelayFrame / ld a,[wTextDelayFrames] / ret
t1 = b'\xCD' + le(DelayFrame) + b'\xFA' + le(wTextDelayFrames) + b'\xC9'
# T2: ld a,[w2DMenuFlags1] / bit 6,a / call z,DelayFrame / ld a,[w2DMenuFlags1] / ret
#     (menus with sprite animation already halt in PlaySpriteAnimationsAndDelayFrame)
t2 = (b'\xFA' + le(w2DMenuFlags1) + b'\xCB\x77' + b'\xCC' + le(DelayFrame)
      + b'\xFA' + le(w2DMenuFlags1) + b'\xC9')
tramp = t1 + t2

# free ROM0 space from the .map, staying clear of $3FFE-$3FFF
free = None
in_rom0 = False
for line in open(MAP, encoding='utf-8', errors='replace'):
    if line.startswith('ROM0 bank #0'):
        in_rom0 = True
        continue
    if in_rom0 and re.match(r'^\S', line):
        break
    m = re.search(r'EMPTY: \$([0-9a-f]+)-\$([0-9a-f]+)', line) if in_rom0 else None
    if m:
        s, e = int(m.group(1), 16), min(int(m.group(2), 16), 0x3FFD)
        if s >= 0x100 and e - s + 1 >= len(tramp):
            free = s
if free is None:
    die('no free ROM0 space for trampolines')
if any(rom[free:free + len(tramp)]):
    die('free space at 0x%X is not empty' % free)
T1, T2 = free, free + len(t1)

# ---- patch sites -------------------------------------------------------
o_wait = off(b_wait, a_wait)
expect(o_wait, b'\xFA' + le(wTextDelayFrames) + b'\xA7', 'PrintLetterDelay.wait')

o_menu = off(b_menu, a_menu) + 7    # call UpdateTimeAndPals / call Menu_WasButtonPressed / ret c
expect(o_menu - 1, b'\xD8\xFA' + le(w2DMenuFlags1) + b'\xCB\x7F', 'Do2DMenuRTCJoypad.loopRTC')

o_rel = rom.find(b'\x3E\x09\xCD' + le(SkipMusic), off(b_rel, a_rel), off(b_rel, a_rel) + 0x60)
if o_rel < 0:
    die('ld a, 9 / call SkipMusic not found in ReloadTilesetAndPalettes')

def hexbytes(b):
    return ' '.join('%02X' % x for x in b)

entries = [
    ('SlowDown Trampolines', T1, tramp),
    ('Msg SlowDown',         o_wait, b'\xCD' + le(T1)),
    ('Msg Select SlowDown',  o_menu, b'\xCD' + le(T2)),
    ('FieldRewriteMain SlowDown', o_rel, b'\x3E\x03'),
]

block = [BEGIN,
         '; Busy-wait loops -> halt (DelayFrame), same idea as the Korean VC patch.',
         '; T1 @0x%04X: call DelayFrame / ld a,[wTextDelayFrames] / ret' % T1,
         '; T2 @0x%04X: menu wait, DelayFrame only if w2DMenuFlags1 bit 6 is clear' % T2,
         '']
for name, addr, data in entries:
    block += ['[%s]' % name, 'Mode = 1', 'Address = 0x%X' % addr,
              'Fixcode = a%d:%s' % (len(data), hexbytes(data)), '']
block.append(END)

# ---- write patch (keep UTF-8 BOM + CRLF) ------------------------------
raw = open(PATCH, 'rb').read()
bom = raw.startswith(b'\xef\xbb\xbf')
text = raw.decode('utf-8-sig')
text = re.sub(r'\r?\n?' + re.escape(BEGIN) + r'.*?' + re.escape(END) + r'[^\n]*', '', text, flags=re.S)
text = text.rstrip('\r\n') + '\r\n\r\n' + '\r\n'.join(block) + '\r\n'
open(PATCH, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + text.encode('utf-8'))

for name, addr, data in entries:
    print('%-26s 0x%06X  %s' % (name, addr, hexbytes(data)))
print('vc_slowdown: updated', PATCH)

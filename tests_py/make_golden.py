import sys, os, types

root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
stub = types.ModuleType("tkinterdnd2")
stub.DND_FILES = "<<Drag>>"
class _TD:  # noqa: only attribute access is imported at module level
    pass
stub.TkinterDnD = _TD
sys.modules["tkinterdnd2"] = stub
sys.path.insert(0, root)
import isel_to_gcode_gui as G
G._NUMPY = False  # canonical for the synthetic fixture: pure == numpy bytes here (verified);
# the shipped-path goldens are Mert's real programs in demo*_expected.*.ngc, compared in real_programs.rs

fixture = os.path.join(root, "tests", "fixtures", "sample.isel")
outdir = os.path.join(root, "tests", "fixtures")
arc = "--arcs" in sys.argv
tag = "arcs" if arc else "g1"
out_ngc = os.path.join(outdir, "sample_expected.%s.ngc" % tag)

logs = []
t = G.convert_file(fixture, out_ngc, lambda m: logs.append(m),
                   lambda p, txt: None, use_arc_commands=arc)
with open(os.path.join(outdir, "sample_expected.%s.log" % tag), "w",
          encoding="utf-8", newline="\n") as f:
    f.write("\n".join(logs) + "\n")
with open(os.path.join(outdir, "sample_expected.%s.time" % tag), "w") as f:
    f.write(repr(t))
print("GOLDEN OK", tag, "ngc_bytes=", os.path.getsize(out_ngc),
      "time=", repr(t))

# Generates the small CFF test fonts used by PdfInfrastructureTests (requires fontTools).
#   python make_test_fonts.py
# OysterTestCff.otf    : name-keyed CFF with global and local subroutines
# OysterTestCffCid.otf : CID-keyed CFF (ROS Adobe-Japan1) with two FDs, each with local subroutines
from fontTools.cffLib import FDArrayIndex, FDSelect, FontDict, GlobalSubrsIndex, PrivateDict, SubrsIndex
from fontTools.fontBuilder import FontBuilder
from fontTools.misc.psCharStrings import T2CharString

GLYPHS = [".notdef", "space", "A", "B", "uni65E5", "uni672C"]
CMAP = {0x20: "space", 0x41: "A", 0x42: "B", 0x65E5: "uni65E5", 0x672C: "uni672C"}
ADVANCES = {".notdef": 500, "space": 250, "A": 600, "B": 620, "uni65E5": 1000, "uni672C": 1000}

# Subroutine bias is 107 for fewer than 1240 subrs
BIAS = 107


def box(x0, y0, x1, y1):
    return [x0, y0, "rmoveto", x1 - x0, "hlineto", y1 - y0, "vlineto", x0 - x1, "hlineto"]


def build(path, cid_keyed):
    fb = FontBuilder(1000, isTTF=False)
    fb.setupGlyphOrder(GLYPHS)
    fb.setupCharacterMap(CMAP)

    global_subrs = GlobalSubrsIndex()
    for program in (box(100, 0, 500, 100) + ["return"], box(100, 600, 500, 700) + ["return"], ["return"]):
        global_subrs.append(T2CharString(program=program))

    def local_subrs():
        subrs = SubrsIndex()
        for program in (box(200, 200, 400, 400) + ["return"], box(50, 50, 80, 80) + ["return"], ["return"]):
            subrs.append(T2CharString(program=program))
        return subrs

    programs = {
        ".notdef": [500] + box(50, 0, 450, 700) + ["endchar"],
        "space": [250, "endchar"],
        "A": [600, 0 - BIAS, "callgsubr", 0 - BIAS, "callsubr", "endchar"],
        "B": [620, 1 - BIAS, "callgsubr", "endchar"],
        "uni65E5": [1000] + box(100, -50, 900, 800) + [0 - BIAS, "callsubr", "endchar"],
        "uni672C": [1000, 1 - BIAS, "callsubr", 0 - BIAS, "callgsubr", "endchar"],
    }

    charstrings = {name: T2CharString(program=program) for name, program in programs.items()}
    fb.setupCFF("OysterTestCff", {"FullName": "Oyster Test CFF"}, charstrings, {})
    fb.setupHorizontalMetrics({name: (ADVANCES[name], 0) for name in GLYPHS})
    fb.setupHorizontalHeader(ascent=880, descent=-120)
    fb.setupNameTable({"familyName": "Oyster Test CFF" + (" CID" if cid_keyed else ""), "styleName": "Regular"})
    fb.setupOS2(sTypoAscender=880, sTypoDescender=-120, usWinAscent=880, usWinDescent=120)
    fb.setupPost()

    cff = fb.font["CFF "].cff
    top = cff.topDictIndex[0]
    cff.GlobalSubrs = global_subrs
    top.GlobalSubrs = global_subrs

    if not cid_keyed:
        top.Private.Subrs = local_subrs()
        for cs in top.CharStrings.values():
            cs.private = top.Private
            cs.globalSubrs = global_subrs
    else:
        top.ROS = ("Adobe", "Japan1", 6)
        top.CIDCount = len(GLYPHS)
        fd_array = FDArrayIndex()
        for _ in range(2):
            fd = FontDict()
            fd.setCFF2(False)
            private = PrivateDict()
            private.Subrs = local_subrs()
            private.defaultWidthX = 0
            private.nominalWidthX = 0
            fd.Private = private
            fd.FontMatrix = [0.001, 0, 0, 0.001, 0, 0]
            fd_array.append(fd)
        top.FDArray = fd_array
        select = FDSelect()
        select.format = 3
        select.gidArray = [0, 0, 0, 1, 1, 1]
        top.FDSelect = select
        top.rawDict.pop("Private", None)
        if "Private" in top.__dict__:
            del top.__dict__["Private"]
        # CID-keyed charsets name glyphs cidNNNNN; use non-identity CIDs to exercise the mapping
        cids = [0, 1, 34, 35, 2911, 3613]
        cff.topDictIndex[0].charset = [".notdef"] + ["cid%05d" % c for c in cids[1:]]
        rename = dict(zip(GLYPHS, top.charset))
        fb.font.setGlyphOrder(top.charset)
        top.CharStrings.charStrings = {rename[k]: v for k, v in top.CharStrings.charStrings.items()}
        for index, name in enumerate(top.charset):
            cs = top.CharStrings[name]
            cs.private = fd_array[select.gidArray[index]].Private
            cs.globalSubrs = global_subrs
        fb.font["hmtx"].metrics = {rename[k]: v for k, v in fb.font["hmtx"].metrics.items()}
        cmap = fb.font["cmap"]
        for table in cmap.tables:
            table.cmap = {code: rename[name] for code, name in table.cmap.items()}

    fb.save(path)


build("OysterTestCff.otf", cid_keyed=False)
build("OysterTestCffCid.otf", cid_keyed=True)

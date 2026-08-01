"""Validate all manual examples from VBA_LIB_User_Manual.md against Python references.

Usage: python tests/build_manual_examples.py

Validates 224 hand-curated examples from all 14 chapters of the user manual.
Each example matches a VBA code block or UDF formula in the manual.
"""

import os, sys
import numpy as np
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from tests.test_utils import (
    ensure_excel, teardown, create_workbook, inject_testrunner,
    run_macro, com_to_numpy, write_range, SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER,
)
from tests.crossval.build_common import _flatten_com_result

# Helpers
def _to_com_arg(arg):
    """datetime/date → ISO date string (VBA CDate parses unambiguously)."""
    if isinstance(arg, (datetime, date)):
        return arg.strftime("%Y-%m-%d")
    return arg

def _flatten_list(val):
    """Recursively flatten nested lists/tuples to a flat list."""
    if isinstance(val, (list, tuple)):
        result = []
        for item in val:
            result.extend(_flatten_list(item))
        return result
    return [val]

def _dt(y,m,d): return datetime(y,m,d)
def _easter(y):
    a=y%19;b=y//100;c=y%100;d=b//4;e=b%4;f=(b+8)//25;g=(b-f+1)//3
    h=(19*a+b-d-g+15)%30;i=c//4;k=c%4
    L=(32+2*e+2*i-h-k)%7;m=(a+11*h+22*L)//451
    return date(y,(h+L-7*m+114)//31,((h+L-7*m+114)%31)+1)

def _py_anova_fstat(data):
    """Compute ANOVA F-statistic from 2D data [[group, value], ...]."""
    grp_vals = {}
    for r in data:
        grp_vals.setdefault(r[0], []).append(float(r[1]))
    groups = list(grp_vals.values())
    if len(groups) < 2: return float("nan")
    try:
        from scipy import stats
        return float(stats.f_oneway(*groups).statistic)
    except ImportError:
        return float("nan")

# (chapter, name, module, func, args_fn, py_ref_fn, result_type, tol?)
MANUAL_EXAMPLES = [

    # ======================================================================
    # Chapter 1: ArrayUtils (41 examples)
    # ======================================================================
    ("1","ArrayDims_1D","ArrayUtils","ArrayDims",lambda:([1,2,3],),lambda a:1,"scalar"),
    ("1","ArrayDims_Empty","ArrayUtils","ArrayDims",lambda:(None,),lambda a:0,"scalar"),
    ("1","IsArray1D_true","ArrayUtils","IsArray1D",lambda:([1,2],),lambda a:True,"bool"),
    ("1","IsArray1D_scalar","ArrayUtils","IsArray1D",lambda:("not_array",),lambda a:False,"bool"),
    ("1","ArraySort_asc","ArrayUtils","ArraySort",lambda:([3,1,4,2],True),lambda a:sorted(a[0]),"array"),
    ("1","ArraySort_desc","ArrayUtils","ArraySort",lambda:([3,1,4,2],False),lambda a:sorted(a[0],reverse=True),"array"),
    ("1","ArrayUnique","ArrayUtils","ArrayUnique",lambda:([1,2,2,3,1],),lambda a:[1,2,3],"array"),
    ("1","ArrayFilter_gt","ArrayUtils","ArrayFilterByValue",lambda:([3,1,4,1,5],3,">"),lambda a:[4,5],"array"),
    ("1","ArrayFilter_contains","ArrayUtils","ArrayFilterByValue",lambda:(["ab","cd","ef"],"c","contains"),lambda a:["cd"],"array"),
    ("1","ArrayFilter_no_match","ArrayUtils","ArrayFilterByValue",lambda:([1,2,3],9,">"),lambda a:[],"array"),
    ("1","ArrayCountIf_gt80","ArrayUtils","ArrayCountIf",lambda:([85,92,78,65,91],80,">"),lambda a:3,"scalar"),
    ("1","ArrayCountIf_eq","ArrayUtils","ArrayCountIf",lambda:([1,2,1,3],1),lambda a:2,"scalar"),
    ("1","ArrayMin","ArrayUtils","ArrayMin",lambda:([85,92,78,65,91],),lambda a:65,"scalar"),
    ("1","ArrayMax","ArrayUtils","ArrayMax",lambda:([85,92,78,65,91],),lambda a:92,"scalar"),
    ("1","ArraySum","ArrayUtils","ArraySum",lambda:([85,92,78,65,91],),lambda a:411,"scalar"),
    ("1","ArrayConcat_arrs","ArrayUtils","ArrayConcat",lambda:([1,2],[3,4]),lambda a:[1,2,3,4],"array"),
    ("1","ArrayConcat_scalar_arr","ArrayUtils","ArrayConcat",lambda:(42,[1,2]),lambda a:[42,1,2],"array"),
    ("1","ArrayConcat_arr_scalar","ArrayUtils","ArrayConcat",lambda:([1,2],99),lambda a:[1,2,99],"array"),
    ("1","ArraySlice_mid","ArrayUtils","ArraySlice",lambda:([10,20,30,40],1,2),lambda a:[20,30],"array"),
    ("1","ArraySlice_neg","ArrayUtils","ArraySlice",lambda:([10,20,30,40],-1,1),lambda a:[40],"array"),
    ("1","ArraySlice_default","ArrayUtils","ArraySlice",lambda:([10,20,30,40],0,-1),lambda a:[10,20,30,40],"array"),
    ("1","ArrayFlatten","ArrayUtils","ArrayFlatten",lambda:([[1,2],[3,4],[5,6]],),lambda a:[1,2,3,4,5,6],"array"),
    ("1","ArrayTranspose1D_col","ArrayUtils","ArrayTranspose1D",lambda:([1,2,3],True),lambda a:[[1],[2],[3]],"array"),
    ("1","ArrayFind_found","ArrayUtils","ArrayFind",lambda:([10,20,30,40],30),lambda a:2,"scalar"),
    ("1","ArrayFind_not_found","ArrayUtils","ArrayFind",lambda:([10,20,30],99),lambda a:-1,"scalar"),
    ("1","ArrayContains_true","ArrayUtils","ArrayContains",lambda:(["苹果","香蕉","橙子"],"香蕉"),lambda a:True,"bool"),
    ("1","ArrayContains_false","ArrayUtils","ArrayContains",lambda:([1,2,3],4),lambda a:False,"bool"),
    ("1","ArrayEqual_true","ArrayUtils","ArrayEqual",lambda:([1,2,3],[1,2,3]),lambda a:True,"bool"),
    ("1","ArrayEqual_false_len","ArrayUtils","ArrayEqual",lambda:([1,2],[1,2,3]),lambda a:False,"bool"),
    ("1","ArrayProduct","ArrayUtils","ArrayProduct",lambda:([2,3,4],),lambda a:24,"scalar"),
    ("1","CumSum","ArrayUtils","CumSum",lambda:([3,1,4],),lambda a:[3,4,8],"array"),
    ("1","ArgSort","ArrayUtils","ArgSort",lambda:([30,10,20],True),lambda a:[1,2,0],"array"),
    ("1","ArrayAny_true","ArrayUtils","ArrayAny",lambda:([1,2,10],5,">"),lambda a:True,"bool"),
    ("1","ArrayAny_false","ArrayUtils","ArrayAny",lambda:([1,2,3],5,">"),lambda a:False,"bool"),
    ("1","ArrayAll_true","ArrayUtils","ArrayAll",lambda:([2,4,6],1,">"),lambda a:True,"bool"),
    ("1","ArrayAll_false","ArrayUtils","ArrayAll",lambda:([1,2,6],1,">"),lambda a:False,"bool"),
    ("1","LinSpace","ArrayUtils","LinSpace",lambda:(0.,1.,5),lambda a:[0.,.25,.5,.75,1.],"array",1e-8),
    ("1","RangeFill","ArrayUtils","RangeFill",lambda:(0.,5,2.),lambda a:[0.,2.,4.,6.,8.],"array",1e-8),
    ("1","ArrayToString","ArrayUtils","ArrayToString",lambda:([1,2,3],"-"),lambda a:"1-2-3","string"),
    ("1","ArrayReverse","ArrayUtils","ArrayReverse",lambda:([1,2,3],),lambda a:[3,2,1],"array"),
    ("1","ArrayChunk","ArrayUtils","ArrayChunk",lambda:([1,2,3,4,5],2),lambda a:[[1,2],[3,4],[5,None]],"array"),

    # ======================================================================
    # Chapter 2: DictSetUtils (10 examples)
    # ======================================================================
    ("2","SetUnion","DictSetUtils","SetUnion",lambda:([1,2,3],[3,4,5]),lambda a:[1,2,3,4,5],"array"),
    ("2","SetIntersect","DictSetUtils","SetIntersect",lambda:([1,2,3],[2,3,4]),lambda a:[2,3],"array"),
    ("2","SetDifference","DictSetUtils","SetDifference",lambda:([1,2,3],[2,4]),lambda a:[1,3],"array"),
    ("2","SetSymDiff","DictSetUtils","SetSymDifference",lambda:([1,2,3],[2,3,4]),lambda a:[1,4],"array"),
    ("2","SetSymDiff_manual","DictSetUtils","SetSymDifference",lambda:([1,2,3],[2,4]),lambda a:[1,3,4],"array"),
    ("2","SetIsSubset_true","DictSetUtils","SetIsSubset",lambda:([1,2],[1,2,3]),lambda a:True,"bool"),
    ("2","SetIsSubset_false","DictSetUtils","SetIsSubset",lambda:([1,4],[1,2,3]),lambda a:False,"bool"),
    ("2","SetEqual_true","DictSetUtils","SetEqual",lambda:([1,2,3],[3,2,1]),lambda a:True,"bool"),
    ("2","SetEqual_false","DictSetUtils","SetEqual",lambda:([1,2],[1,2,3]),lambda a:False,"bool"),
    ("2","GroupCount","DictSetUtils","GroupCount",lambda:(["A","B","A","C","B","A"],),lambda a:[["A",3],["B",2],["C",1]],"array"),

    # ======================================================================
    # Chapter 3: PivotUtils (4 examples)
    # ======================================================================
    ("3","GroupBy_SUM","PivotUtils","GroupBy",
     lambda:([["ID","Name","Score"],[1,"A",100],[1,"B",150],[2,"A",80],[2,"B",120]],1,3,"SUM"),
     lambda a:_py_grp(a[0][1:],a[1],a[2],a[3]),"array",1e-10,True),
    ("3","GroupBy_AVG","PivotUtils","GroupBy",
     lambda:([["ID","Name","Score"],[1,"A",100],[1,"B",150],[2,"A",80],[2,"B",120]],1,3,"AVG"),
     lambda a:_py_grp(a[0][1:],a[1],a[2],a[3]),"array",1e-10,True),
    ("3","GroupBy_COUNT","PivotUtils","GroupBy",
     lambda:([["ID","Name","Score"],[1,"A",100],[1,"B",150],[2,"A",80]],1,3,"COUNT"),
     lambda a:_py_grp(a[0][1:],a[1],a[2],a[3]),"array",1e-10,True),
    ("3","CrossJoin","PivotUtils","CrossJoin",
     lambda:(["Name","张三","李四"],["Subject","数学","英语"]),
     lambda a:[[x,y] for x in a[0][1:] for y in a[1][1:]],"array",1e-10,True),

    # ======================================================================
    # Chapter 4: SqlUtils (1 example — ADODB needs saved workbook)
    # ======================================================================
    ("4","SqlListSheets","SqlUtils","SqlListSheets",
     lambda:(),lambda a:["TestResults","TestData"],"array",1e-10,False,
     True, "workbook-dependent sheet list; manual validation sufficient"),

    # ======================================================================
    # Chapter 5: LinearUtils (5 examples)
    # ======================================================================
    ("5","MatrixMultiply","LinearUtils","MatrixMultiply",
     lambda:([[1,2],[3,4]],[[1,0],[0,1]]),
     lambda a:np.array(a[0]),"array",1e-10),
    ("5","MatrixDeterminant","LinearUtils","MatrixDeterminant",
     lambda:([[1,2],[3,4]],),
     lambda a:float(np.linalg.det([[1.,2.],[3.,4.]])),"scalar",1e-8),
    ("5","MatrixTranspose","LinearUtils","MatrixTranspose",
     lambda:([[1,2],[3,4],[5,6]],),
     lambda a:np.array([[1,2],[3,4],[5,6]]).T,"array",1e-10),
    ("5","IdentityMatrix","LinearUtils","IdentityMatrix",
     lambda:(3,),lambda a:np.eye(3),"array",1e-10),
    ("5","VectorDot","LinearUtils","VectorDot",
     lambda:([1.,2.,3.],[4.,5.,6.]),
     lambda a:float(np.dot([1.,2.,3.],[4.,5.,6.])),"scalar",1e-10),

    # ======================================================================
    # Chapter 6: StatsUtils (24 examples)
    # ======================================================================
    ("6","Mean","StatsUtils","Mean",lambda:([1.,2.,3.,4.,5.],),lambda a:3.,"scalar"),
    ("6","Median_odd","StatsUtils","Median",lambda:([1.,2.,3.,4.,100.],),lambda a:3.,"scalar"),
    ("6","Median_even","StatsUtils","Median",lambda:([1.,2.,3.,4.],),lambda a:2.5,"scalar"),
    ("6","StdDev","StatsUtils","StdDev",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.std([1.,2.,3.,4.,5.],ddof=1)),"scalar",1e-8),
    ("6","Variance","StatsUtils","Variance",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.var([1.,2.,3.,4.,5.],ddof=1)),"scalar",1e-8),
    ("6","Mode","StatsUtils","Mode",lambda:([1.,2.,2.,3.,4.],),lambda a:2.,"scalar"),
    ("6","Min","StatsUtils","Min",lambda:([5.,3.,1.,4.,2.],),lambda a:1.,"scalar"),
    ("6","Max","StatsUtils","Max",lambda:([5.,3.,1.,4.,2.],),lambda a:5.,"scalar"),
    ("6","GeometricMean","StatsUtils","GeometricMean",lambda:([1.05,1.08,1.03,1.06],),lambda a:float(np.exp(np.mean(np.log([1.05,1.08,1.03,1.06])))),"scalar",1e-6),
    ("6","HarmonicMean","StatsUtils","HarmonicMean",lambda:([2.,3.,6.],),lambda a:3.,"scalar",1e-6),
    ("6","RootMeanSquare","StatsUtils","RootMeanSquare",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.sqrt(np.mean(np.square([1.,2.,3.,4.,5.])))),"scalar",1e-8),
    ("6","Percentile_median","StatsUtils","Percentile",lambda:([1.,2.,3.,4.,5.],0.5),lambda a:3.,"scalar"),
    ("6","IQR","StatsUtils","IQR",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.percentile([1.,2.,3.,4.,5.],75)-np.percentile([1.,2.,3.,4.,5.],25)),"scalar",1e-8),
    ("6","StandardError","StatsUtils","StandardError",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.std([1.,2.,3.,4.,5.],ddof=1)/np.sqrt(5)),"scalar",1e-8),
    ("6","ZScore","StatsUtils","ZScore",lambda:([1.,2.,3.,4.,5.],1.),lambda a:(1.-np.mean([1.,2.,3.,4.,5.]))/np.std([1.,2.,3.,4.,5.],ddof=1),"scalar",1e-6),
    ("6","WeightedMean","StatsUtils","WeightedMean",lambda:([1.,2.,3.],[1.,1.,2.]),lambda a:float(np.average([1.,2.,3.],weights=[1.,1.,2.])),"scalar",1e-8),
    ("6","Correlation_pos","StatsUtils","Correlation",lambda:([1,2,3,4,5],[2,4,6,8,10]),lambda a:1.,"scalar",1e-8),
    ("6","Correlation_neg","StatsUtils","Correlation",lambda:([1,2,3,4,5],[10,8,6,4,2]),lambda a:-1.,"scalar",1e-8),
    ("6","Covariance","StatsUtils","Covariance",lambda:([1,2,3,4,5],[2,4,6,8,10]),lambda a:float(np.cov([1,2,3,4,5],[2,4,6,8,10],ddof=1)[0,1]),"scalar",1e-6),
    ("6","RSquare","StatsUtils","RSquare",lambda:([2,4,6,8,10],[2,4,6,8,10]),lambda a:1.,"scalar",1e-8),
    ("6","StdDevP","StatsUtils","StdDevP",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.std([1.,2.,3.,4.,5.],ddof=0)),"scalar",1e-8),
    ("6","VarianceP","StatsUtils","VarianceP",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.var([1.,2.,3.,4.,5.],ddof=0)),"scalar",1e-8),
    ("6","LinInterp","StatsUtils","LinInterp",lambda:(2.5,[1.,2.,3.],[10.,20.,30.]),lambda a:float(np.interp(2.5,[1.,2.,3.],[10.,20.,30.])),"scalar",1e-8),
    ("6","MeanAbsDev","StatsUtils","MeanAbsDev",lambda:([1.,2.,3.,4.,5.],),lambda a:float(np.mean(np.abs(np.array([1.,2.,3.,4.,5.])-3.))),"scalar",1e-8),

    # ======================================================================
    # Chapter 7: RegressUtils (1 example)
    # ======================================================================
    ("7","ANOVAOneWay_Fstat","RegressUtils","ANOVAOneWay_Fstat",
     lambda:([[1,23],[1,25],[1,22],[2,30],[2,32],[2,31],[3,28],[3,27],[3,29]],1,2,False),
     lambda a:_py_anova_fstat(a[0]),"scalar",0.1),

    # ======================================================================
    # Chapter 8: StringUtils (49 examples)
    # ======================================================================
    ("8","RemoveChars","StringUtils","RemoveChars",lambda:("AB-12/CD-34","-/"),lambda a:"AB12CD34","string"),
    ("8","KeepChars","StringUtils","KeepChars",lambda:("Tel: 555-1234","0123456789-"),lambda a:"555-1234","string"),
    ("8","KeepChars_digits","StringUtils","KeepChars",lambda:("Tel: 555-1234","0123456789"),lambda a:"5551234","string"),
    ("8","NormalizeWs","StringUtils","NormalizeWhitespace",lambda:("  Hello   World  ",),lambda a:"Hello World","string"),
    ("8","NormalizeWs_tabs","StringUtils","NormalizeWhitespace",lambda:("a\tb\tc",),lambda a:"a b c","string"),
    ("8","ToTitleCase","StringUtils","ToTitleCase",lambda:("hello world",),lambda a:"Hello World","string"),
    ("8","RemoveDiacritics","StringUtils","RemoveDiacritics",lambda:("crème brûlée",),lambda a:"creme brulee","string"),
    ("8","Slugify","StringUtils","Slugify",lambda:("Hello World!",),lambda a:"hello-world","string"),
    ("8","Base64Encode","StringUtils","Base64Encode",lambda:("Hello",),lambda a:"SGVsbG8=","string"),
    ("8","Base64Encode_empty","StringUtils","Base64Encode",lambda:("",),lambda a:"","string"),
    ("8","Base64Decode","StringUtils","Base64Decode",lambda:("SGVsbG8=",),lambda a:"Hello","string"),
    ("8","URLEncode","StringUtils","URLEncode",lambda:("Hello World",),lambda a:"Hello%20World","string"),
    ("8","URLDecode","StringUtils","URLDecode",lambda:("Hello%20World",),lambda a:"Hello World","string"),
    ("8","HTMLEncode","StringUtils","HTMLEncode",lambda:("<div>",),lambda a:"&lt;div&gt;","string"),
    ("8","HTMLEncode_apos","StringUtils","HTMLEncode",lambda:("<div class='a'>",),lambda a:"&lt;div class=&#39;a&#39;&gt;","string"),
    ("8","HTMLDecode","StringUtils","HTMLDecode",lambda:("&lt;p&gt;Hi&lt;/p&gt;",),lambda a:"<p>Hi</p>","string"),
    ("8","Levenshtein_kitten","StringUtils","LevenshteinDistance",lambda:("kitten","sitting"),lambda a:3,"scalar"),
    ("8","Levenshtein_same","StringUtils","LevenshteinDistance",lambda:("abc","abc"),lambda a:0,"scalar"),
    ("8","Levenshtein_empty","StringUtils","LevenshteinDistance",lambda:("","hello"),lambda a:5,"scalar"),
    ("8","Soundex_Robert","StringUtils","Soundex",lambda:("Robert",),lambda a:"R163","string"),
    ("8","Soundex_Rupert","StringUtils","Soundex",lambda:("Rupert",),lambda a:"R163","string"),
    ("8","ExtractBetween","StringUtils","ExtractBetween",lambda:("<title>Hello</title>","<title>","</title>"),lambda a:"Hello","string"),
    ("8","ReverseStr","StringUtils","ReverseString",lambda:("ABC",),lambda a:"CBA","string"),
    ("8","ReverseStr_empty","StringUtils","ReverseString",lambda:("",),lambda a:"","string"),
    ("8","CountSubstr","StringUtils","CountSubstring",lambda:("banana","an"),lambda a:2,"scalar"),
    ("8","StartsWith","StringUtils","StartsWith",lambda:("Hello World","Hello"),lambda a:True,"bool"),
    ("8","EndsWith","StringUtils","EndsWith",lambda:("report.xlsx",".xlsx"),lambda a:True,"bool"),
    ("8","LeftOf","StringUtils","LeftOf",lambda:("john.doe@example.com","@"),lambda a:"john.doe","string"),
    ("8","RightOf","StringUtils","RightOf",lambda:("C:\\Data\\file.txt","\\"),lambda a:"Data\\file.txt","string"),
    ("8","IsEmail","StringUtils","IsEmail",lambda:("user@example.com",),lambda a:True,"bool"),
    ("8","IsEmail_invalid","StringUtils","IsEmail",lambda:("invalid",),lambda a:False,"bool"),
    ("8","IsUrl","StringUtils","IsUrl",lambda:("https://example.com",),lambda a:True,"bool"),
    ("8","IsUrl_invalid","StringUtils","IsUrl",lambda:("not-a-url",),lambda a:False,"bool"),
    ("8","Coalesce","StringUtils","Coalesce",lambda:(None,"","Hello","World"),lambda a:"Hello","string"),
    ("8","Coalesce_none","StringUtils","Coalesce",lambda:(None,"",None),lambda a:"","string"),
    ("8","Repeat","StringUtils","Repeat",lambda:("ab",3),lambda a:"ababab","string"),
    ("8","Repeat_zero","StringUtils","Repeat",lambda:("ab",0),lambda a:"","string"),
    ("8","Truncate","StringUtils","Truncate",lambda:("Hello World",8),lambda a:"Hello...","string"),
    ("8","PadLeft","StringUtils","PadLeft",lambda:("42",5,"0"),lambda a:"00042","string"),
    ("8","PadRight","StringUtils","PadRight",lambda:("Name",8),lambda a:"Name    ","string"),
    ("8","IsNullOrEmpty_true","StringUtils","IsNullOrEmpty",lambda:("",),lambda a:True,"bool"),
    ("8","IsNullOrEmpty_false","StringUtils","IsNullOrEmpty",lambda:(" ",),lambda a:False,"bool"),
    ("8","IsNullOrWs_true","StringUtils","IsNullOrWhitespace",lambda:("  ",),lambda a:True,"bool"),
    ("8","IsNullOrWs_false","StringUtils","IsNullOrWhitespace",lambda:("a",),lambda a:False,"bool"),
    ("8","TextJoin","StringUtils","TextJoin",lambda:(", ",["A","B","C"]),lambda a:"A, B, C","string"),
    ("8","NthWord","StringUtils","NthWord",lambda:("apple,banana,cherry",2,","),lambda a:"banana","string"),
    ("8","CommonPrefix","StringUtils","CommonPrefix",lambda:("flower","flow"),lambda a:"flow","string"),
    ("8","CommonPrefix_fl","StringUtils","CommonPrefix",lambda:("flower","flight"),lambda a:"fl","string"),
    ("8","CommonPrefix_case","StringUtils","CommonPrefix",lambda:("abc","ABC"),lambda a:"abc","string"),

    # ======================================================================
    # Chapter 9: RegexUtils (11 examples)
    # ======================================================================
    ("9","RegexIsMatch_yes","RegexUtils","RegexIsMatch",lambda:("abc123",r"\d+"),lambda a:True,"bool"),
    ("9","RegexIsMatch_no","RegexUtils","RegexIsMatch",lambda:("abc",r"^\d+$"),lambda a:False,"bool"),
    ("9","RegexExtract","RegexUtils","RegexExtract",lambda:("Phone: 555-1234",r"\d{3}-\d{4}"),lambda a:"555-1234","string"),
    ("9","RegexExtractAll","RegexUtils","RegexExtractAll",lambda:("a1 b2 c3",r"\w\d"),lambda a:["a1","b2","c3"],"array"),
    ("9","RegexExtractGroups","RegexUtils","RegexExtractGroups",lambda:("Name: John, Age: 30",r"(\w+): (\w+)"),lambda a:[["Name","John"],["Age","30"]],"array"),
    ("9","RegexIsFullMatch_yes","RegexUtils","RegexIsFullMatch",lambda:("12345",r"\d+"),lambda a:True,"bool"),
    ("9","RegexIsFullMatch_no","RegexUtils","RegexIsFullMatch",lambda:("abc123",r"\d+"),lambda a:False,"bool"),
    ("9","RegexReplace","RegexUtils","RegexReplace",lambda:("2024-01-15",r"(\d{4})-(\d{2})-(\d{2})","$3/$2/$1"),lambda a:"15/01/2024","string"),
    ("9","RegexSplit","RegexUtils","RegexSplit",lambda:("a,b; c|d",r"[,;|]\s*"),lambda a:["a","b","c","d"],"array"),
    ("9","RegexCount","RegexUtils","RegexCount",lambda:("The fat cat sat on the mat",r"\b\w{3}\b"),lambda a:6,"scalar"),
    ("9","RegexEscape","RegexUtils","RegexEscape",lambda:("1+1=2?",),lambda a:r"1\+1=2\?","string"),

    # ======================================================================
    # Chapter 10: JsonUtils (7 examples)
    # ======================================================================
    ("10","JsonIsValid_obj","JsonUtils","JsonIsValid",lambda:('{"a":1}',),lambda a:True,"bool"),
    ("10","JsonIsValid_string","JsonUtils","JsonIsValid",lambda:('"hello"',),lambda a:True,"bool"),
    ("10","JsonIsValid_number","JsonUtils","JsonIsValid",lambda:("42",),lambda a:True,"bool"),
    ("10","JsonIsValid_bad","JsonUtils","JsonIsValid",lambda:("{bad}",),lambda a:False,"bool"),
    ("10","JsonIsValid_empty","JsonUtils","JsonIsValid",lambda:("",),lambda a:False,"bool"),
    ("10","JsonStringify","JsonUtils","JsonStringify",lambda:([1,2,3],),lambda a:"[1,2,3]","string"),
    ("10","JsonStringify_empty","JsonUtils","JsonStringify",lambda:([],),lambda a:"[]","string"),

    # ======================================================================
    # Chapter 11: XmlUtils (6 examples)
    # ======================================================================
    ("11","XmlValidate_valid","XmlUtils","XmlValidate",lambda:("<root/>",),lambda a:True,"bool"),
    ("11","XmlValidate_invalid","XmlUtils","XmlValidate",lambda:("<root><bad>",),lambda a:False,"bool"),
    ("11","XmlGet_path","XmlUtils","XmlGet",lambda:("<a><b>hello</b></a>","/a/b"),lambda a:"hello","string"),
    ("11","XmlGetAttr_id","XmlUtils","XmlGetAttr",lambda:("<user id='007'/>","/user","id"),lambda a:"007","string"),
    ("11","XmlToRange_basic","XmlUtils","XmlToRange",lambda:("<rows><row><a>1</a><b>10</b></row></rows>","/rows/row",["a","b"]),lambda a:[[1.,10.]],"array",1e-10),
    ("11","UDF_XML_VALIDATE","XmlUtils","UDF_XML_VALIDATE",lambda:("<root/>",),lambda a:True,"bool"),

    # ======================================================================
    # Chapter 12: DateTimeUtils (22 examples)
    # ======================================================================
    ("12","ISOWeekNum_jan1","DateTimeUtils","ISOWeekNum",lambda:(_dt(2025,1,1),),lambda a:1,"scalar"),
    ("11","ISOWeekNum_dec30","DateTimeUtils","ISOWeekNum",lambda:(_dt(2024,12,30),),lambda a:1,"scalar"),
    ("11","IsLeapYear_2024","DateTimeUtils","IsLeapYear",lambda:(2024,),lambda a:True,"bool"),
    ("11","IsLeapYear_2023","DateTimeUtils","IsLeapYear",lambda:(2023,),lambda a:False,"bool"),
    ("11","IsLeapYear_2000","DateTimeUtils","IsLeapYear",lambda:(2000,),lambda a:True,"bool"),
    ("11","DaysInMonth_feb_leap","DateTimeUtils","DaysInMonth",lambda:(_dt(2024,2,15),),lambda a:29,"scalar"),
    ("11","DaysInMonth_feb_nonleap","DateTimeUtils","DaysInMonth",lambda:(_dt(2023,2,15),),lambda a:28,"scalar"),
    ("11","Quarter_q1","DateTimeUtils","Quarter",lambda:(_dt(2024,2,15),),lambda a:1,"scalar"),
    ("11","Quarter_q3","DateTimeUtils","Quarter",lambda:(_dt(2024,7,15),),lambda a:3,"scalar"),
    ("11","DayOfYear_jan1","DateTimeUtils","DayOfYear",lambda:(_dt(2024,1,1),),lambda a:1,"scalar"),
    ("11","DayOfYear_dec31_leap","DateTimeUtils","DayOfYear",lambda:(_dt(2024,12,31),),lambda a:366,"scalar"),
    ("11","FirstDayOfMonth","DateTimeUtils","FirstDayOfMonth",lambda:(_dt(2024,2,15),),lambda a:_dt(2024,2,1),"scalar"),
    ("11","LastDayOfMonth_feb","DateTimeUtils","LastDayOfMonth",lambda:(_dt(2024,2,15),),lambda a:_dt(2024,2,29),"scalar"),
    ("11","LastDayOfMonth_nonleap","DateTimeUtils","LastDayOfMonth",lambda:(_dt(2023,2,15),),lambda a:_dt(2023,2,28),"scalar"),
    ("11","AgeYears","DateTimeUtils","AgeYears",lambda:(_dt(1990,6,15),_dt(2025,6,15)),lambda a:35,"scalar"),
    ("11","Easter_2025","DateTimeUtils","Easter",lambda:(2025,),lambda a:_easter(2025),"scalar"),
    ("11","Easter_2024","DateTimeUtils","Easter",lambda:(2024,),lambda a:_easter(2024),"scalar"),
    ("11","IsWeekend_sat","DateTimeUtils","IsWeekend",lambda:(_dt(2025,6,14),),lambda a:True,"bool"),
    ("11","IsWeekend_wed","DateTimeUtils","IsWeekend",lambda:(_dt(2025,6,11),),lambda a:False,"bool"),
    ("11","StartOfWeek","DateTimeUtils","StartOfWeek",lambda:(_dt(2025,6,11),2),lambda a:_dt(2025,6,9),"scalar"),
    ("11","DaysInYear_leap","DateTimeUtils","DaysInYear",lambda:(2024,),lambda a:366,"scalar"),
    ("11","DaysInYear_nonleap","DateTimeUtils","DaysInYear",lambda:(2023,),lambda a:365,"scalar"),

    # ======================================================================
    # Chapter 12: RangeUtils (6 examples)
    # ======================================================================
    ("12","ColLetter_A","RangeUtils","ColLetter",lambda:(1,),lambda a:"A","string"),
    ("12","ColLetter_Z","RangeUtils","ColLetter",lambda:(26,),lambda a:"Z","string"),
    ("12","ColLetter_AA","RangeUtils","ColLetter",lambda:(27,),lambda a:"AA","string"),
    ("12","ColNumber_A","RangeUtils","ColNumber",lambda:("A",),lambda a:1,"scalar"),
    ("12","ColNumber_Z","RangeUtils","ColNumber",lambda:("Z",),lambda a:26,"scalar"),
    ("12","GetCellAddress","RangeUtils","GetCellAddress",lambda:(5,3,True),lambda a:"$C$5","string"),

    # ======================================================================
    # Chapter 13: FileSystemUtils (23 examples) -- path-only functions
    # ======================================================================
    # NormalizePath
    ("13","NormalizePath_fwd","FileSystemUtils","NormalizePath",lambda:("C:/Data/file.txt",),lambda a:"C:\\Data\\file.txt","string"),
    ("13","NormalizePath_double","FileSystemUtils","NormalizePath",lambda:("C://dir//sub",),lambda a:"C:\\dir\\sub","string"),
    ("13","NormalizePath_doubleslash","FileSystemUtils","NormalizePath",lambda:("C:/Data//files",),lambda a:"C:\\Data\\files","string"),
    ("13","NormalizePath_mixed","FileSystemUtils","NormalizePath",lambda:("C:\\dir/sub/file",),lambda a:"C:\\dir\\sub\\file","string"),
    ("13","NormalizePath_empty","FileSystemUtils","NormalizePath",lambda:("",),lambda a:"","string"),
    # PathCombine
    ("13","PathCombine_normal","FileSystemUtils","PathCombine",lambda:("C:\\dir","file.txt"),lambda a:"C:\\dir\\file.txt","string"),
    ("13","PathCombine_trailing","FileSystemUtils","PathCombine",lambda:("C:\\dir\\","file.txt"),lambda a:"C:\\dir\\file.txt","string"),
    ("13","PathCombine_empty","FileSystemUtils","PathCombine",lambda:("","file.txt"),lambda a:"file.txt","string"),
    # GetFileName
    ("13","GetFileName_full","FileSystemUtils","GetFileName",lambda:("C:\\Data\\report.xlsx",),lambda a:"report.xlsx","string"),
    ("13","GetFileName_bare","FileSystemUtils","GetFileName",lambda:("file.txt",),lambda a:"file.txt","string"),
    ("13","GetFileName_empty","FileSystemUtils","GetFileName",lambda:("",),lambda a:"","string"),
    # GetBaseName
    ("13","GetBaseName_normal","FileSystemUtils","GetBaseName",lambda:("C:\\dir\\file.txt",),lambda a:"file","string"),
    ("13","GetBaseName_targz","FileSystemUtils","GetBaseName",lambda:("archive.tar.gz",),lambda a:"archive.tar","string"),
    ("13","GetBaseName_dotfile","FileSystemUtils","GetBaseName",lambda:(".gitignore",),lambda a:"","string"),
    # GetExtension
    ("13","GetExtension_txt","FileSystemUtils","GetExtension",lambda:("file.txt",),lambda a:".txt","string"),
    ("13","GetExtension_targz","FileSystemUtils","GetExtension",lambda:("archive.tar.gz",),lambda a:".gz","string"),
    ("13","GetExtension_none","FileSystemUtils","GetExtension",lambda:("Makefile",),lambda a:"","string"),
    # GetFolderPath
    ("13","GetFolderPath_full","FileSystemUtils","GetFolderPath",lambda:("C:\\dir\\file.txt",),lambda a:"C:\\dir","string"),
    ("13","GetFolderPath_root","FileSystemUtils","GetFolderPath",lambda:("C:\\file.txt",),lambda a:"C:\\","string"),
    ("13","GetFolderPath_cwd","FileSystemUtils","GetFolderPath",lambda:("file.txt",),lambda a:"","string"),
    # IsPathValid
    ("13","IsPathValid_good","FileSystemUtils","IsPathValid",lambda:("C:\\Windows",),lambda a:True,"bool"),
    ("13","IsPathValid_empty","FileSystemUtils","IsPathValid",lambda:("",),lambda a:False,"bool"),
    ("13","IsPathValid_bad_lt","FileSystemUtils","IsPathValid",lambda:("file<name>.txt",),lambda a:False,"bool"),

    # ======================================================================
    # Chapter 14: PhyChemUtils (20 examples)
    # ======================================================================
    ("14","MolWeight_H2O","PhyChemUtils","MolecularWeight",lambda:("H2O",),lambda a:18.015,"scalar",0.05),
    ("14","MolWeight_NaCl","PhyChemUtils","MolecularWeight",lambda:("NaCl",),lambda a:58.443,"scalar",0.05),
    ("14","MolWeight_CO2","PhyChemUtils","MolecularWeight",lambda:("CO2",),lambda a:44.009,"scalar",0.05),
    ("14","MolWeight_CaOH2","PhyChemUtils","MolecularWeight",lambda:("Ca(OH)2",),lambda a:74.093,"scalar",0.1),
    ("14","ConvertTemp_CtoF","PhyChemUtils","ConvertTemperature",lambda:(100.,"C","F"),lambda a:212.,"scalar",0.01),
    ("14","ConvertTemp_CtoK","PhyChemUtils","ConvertTemperature",lambda:(0.,"C","K"),lambda a:273.15,"scalar",0.01),
    ("14","ConvertTemp_FtoC","PhyChemUtils","ConvertTemperature",lambda:(32.,"F","C"),lambda a:0.,"scalar",0.01),
    ("14","ConvertVolume_LtomL","PhyChemUtils","ConvertVolume",lambda:(1.,"L","mL"),lambda a:1000.,"scalar",0.01),
    ("14","ConvertPressure_atm","PhyChemUtils","ConvertPressure",lambda:(1.,"atm","kPa"),lambda a:101.325,"scalar",0.01),
    ("14","ConvertMass_g2kg","PhyChemUtils","ConvertMass",lambda:(500.,"g","kg"),lambda a:0.5,"scalar",0.01),
    ("14","ConvertVolume_gal2L","PhyChemUtils","ConvertVolume",lambda:(1.,"gal","L"),lambda a:3.785,"scalar",0.01),
    ("14","MassToMoles_NaCl","PhyChemUtils","MassToMoles",lambda:(58.44,58.44),lambda a:1.,"scalar",0.001),
    ("14","MolesToMass_water","PhyChemUtils","MolesToMass",lambda:(2.,18.015),lambda a:36.03,"scalar",0.01),
    ("14","IdealGasLaw","PhyChemUtils","IdealGasLaw",lambda:(101325.,0.02445,None,298.15),lambda a:1.,"scalar",0.05),
    ("14","IdealGasLaw_std","PhyChemUtils","IdealGasLaw",lambda:(101325.,0.022414,None,273.15),lambda a:1.,"scalar",0.05),
    ("14","Density","PhyChemUtils","Density",lambda:(100.,None,2.5),lambda a:40.,"scalar",0.001),
    ("14","PercentYield_85","PhyChemUtils","PercentYield",lambda:(8.5,10.),lambda a:85.,"scalar",0.01),
    ("14","PercentYield_100","PhyChemUtils","PercentYield",lambda:(10.,10.),lambda a:100.,"scalar",0.01),
    ("14","DilutionSolve","PhyChemUtils","DilutionSolve",lambda:(10.,100.,5.,None),lambda a:200.,"scalar",0.01),
    ("14","DilutionSolve_c1","PhyChemUtils","DilutionSolve",lambda:(10.,50.,None,200.),lambda a:2.5,"scalar",0.01),
]


def _py_grp(data, gc, ac, fn):
    """Python reference for GroupBy (data rows only, no VBA header)."""
    grp = {}
    for r in data:
        k, v = r[gc-1], float(r[ac-1])
        grp.setdefault(k, []).append(v)
    out = []
    for k in sorted(grp, key=str):
        vs = grp[k]
        fu = fn.upper()
        if fu == "SUM": out.append([k, sum(vs)])
        elif fu == "AVG": out.append([k, sum(vs)/len(vs)])
        elif fu == "COUNT": out.append([k, len(vs)])
        elif fu == "MIN": out.append([k, min(vs)])
        elif fu == "MAX": out.append([k, max(vs)])
    return out


def _com_to_scalar(vba_r):
    if vba_r is None: return None
    if isinstance(vba_r,(int,float)): return float(vba_r)
    if isinstance(vba_r,(tuple,list)):
        if len(vba_r)==0: return None
        f=vba_r[0]
        if isinstance(f,(tuple,list)): return float(f[0]) if len(f)>0 else None
        return float(f)
    try: return float(vba_r)
    except (ValueError, TypeError): return None


def main() -> int:
    print("=" * 60)
    print("  Manual Example Validator")
    print("=" * 60)
    print(f"  Examples: {len(MANUAL_EXAMPLES)}")

    needed = sorted(set(ex[2] for ex in MANUAL_EXAMPLES))
    paths = [os.path.join(VBA_CORE_DIR, n + ".cls") for n in VBA_CORE_IMPORT_ORDER]
    for m in needed:
        p = os.path.join(SRC_DIR, f"{m}.bas")
        if os.path.exists(p): paths.append(p)

    excel = ensure_excel()
    wb = None
    try:
        import tempfile
        out = os.path.join(tempfile.gettempdir(), "vba_manual.xlsm")
        wb = create_workbook(excel, out, paths, import_order=VBA_CORE_IMPORT_ORDER)
        inject_testrunner(wb)

        # Ensure TestData sheet for UDF calls
        td = None
        for sh in wb.Sheets:
            if sh.Name == "TestData": td = sh
        if td is None: td = wb.Sheets.Add(); td.Name = "TestData"

        results = []
        for ex in MANUAL_EXAMPLES:
            ch, name, mod, func, args_fn, py_fn, rtype = ex[:7]
            tol = ex[7] if len(ex) > 7 else 1e-10
            is_udf = ex[8] if len(ex) > 8 else False
            skip_if = ex[9] if len(ex) > 9 else False
            skip_reason = ex[10] if len(ex) > 10 else ""
            if skip_if:
                results.append((mod, f"Ch{ch}/{name}", "SKIP", skip_reason))
                print(f"  SKIP Ch{ch}/{name} — {skip_reason}")
                continue
            label = f"Ch{ch}/{name}"
            try:
                args = args_fn()
                args = tuple(_to_com_arg(a) for a in args)
                if is_udf:
                    # Write array args to TestData sheet, pass Range objects
                    td.UsedRange.ClearContents()
                    rng_args = []
                    r_offset = 1
                    for a in args:
                        if isinstance(a, (list, tuple)):
                            arr = np.asarray(a, dtype=object)
                            if arr.ndim == 0: arr = arr.reshape(1, 1)
                            elif arr.ndim == 1: arr = arr.reshape(-1, 1)
                            nr, nc = arr.shape
                            write_range(td, arr, r_offset, 1)
                            rng_args.append(td.Range(td.Cells(r_offset, 1), td.Cells(r_offset + nr - 1, nc)))
                            r_offset = r_offset + nr + 1  # leave a blank row
                        else:
                            rng_args.append(a)
                    vba_r = run_macro(excel, wb, func, *rng_args)
                else:
                    vba_r = run_macro(excel, wb, f"{mod}.{func}", *args)
                py_v = py_fn(args)
                # UDF GroupBy-style: VBA returns header row before data
                if is_udf and isinstance(vba_r, (tuple, list)) and isinstance(py_v, (list, tuple)):
                    if len(vba_r) == len(py_v) + 1:
                        vba_r = vba_r[1:]  # skip header row
                if rtype == "array":
                    # Detect string array: if py_v or vba_r contains non-numeric data
                    is_str = False
                    if isinstance(py_v, (list, tuple)) and len(py_v) > 0:
                        first = py_v[0]
                        if isinstance(first, (list, tuple)) and len(first) > 0:
                            first = first[0]
                        if isinstance(first, str):
                            try: float(first)
                            except (ValueError, TypeError): is_str = True
                    if not is_str and isinstance(vba_r, (tuple, list)) and len(vba_r) > 0:
                        first = vba_r[0]
                        if isinstance(first, (tuple, list)) and len(first) > 0:
                            first = first[0]
                        if isinstance(first, str):
                            try: float(first)
                            except (ValueError, TypeError): is_str = True

                    if is_str:
                        vs = _flatten_com_result(vba_r) if vba_r is not None else ""
                        ps = ",".join(str(x) for x in _flatten_list(py_v))
                        ok = vs.replace(" ","") == ps.replace(" ","")
                        s = "PASS" if ok else "FAIL"
                        d = "" if ok else f"VBA={vs!r} PY={ps!r}"
                        results.append((mod, label, s, d))
                        print(f"  {s:4s} {label}")
                    else:
                        va = com_to_numpy(vba_r)
                        pa = np.asarray(py_v, dtype=float)
                        if pa.ndim == 1 and va.ndim == 2 and va.shape[1] == 1:
                            pa = pa.reshape(-1, 1)
                        if va.size == 0 and pa.size == 0:
                            me, ok = 0.0, True
                        elif va.size == 0 or pa.size == 0:
                            me, ok = float("inf"), False
                        else:
                            try:
                                me = float(np.nanmax(np.abs(np.nan_to_num(va) - np.nan_to_num(pa))))
                                ok = me <= tol
                            except (ValueError, TypeError):
                                # Shape/type mismatch — fall back to string compare
                                ok = str(vba_r) == str(py_v)
                                me = float("inf") if not ok else 0.
                                if not ok:
                                    print(f"  WARN  {label} — shape/type mismatch, str compare used")
                        s = "PASS" if ok else "FAIL"
                        d = "" if ok else f"max|err|={me:.2e}"
                        results.append((mod, label, s, d))
                        print(f"  {s:4s} {label} {'max|err|='+f'{me:.2e}' if ok else d}")
                elif rtype == "bool":
                    try: vb = bool(int(float(vba_r)))
                    except (ValueError, TypeError): vb = str(vba_r).strip().upper() == "TRUE"
                    ok = vb == bool(py_v)
                    s = "PASS" if ok else "FAIL"
                    d = "" if ok else f"VBA={vb} PY={py_v}"
                    results.append((mod, label, s, d))
                    print(f"  {s:4s} {label}")
                elif rtype == "string":
                    vs = str(vba_r).strip() if vba_r is not None else ""
                    ok = vs == str(py_v).strip()
                    s = "PASS" if ok else "FAIL"
                    d = "" if ok else f"VBA={vs!r} PY={py_v!r}"
                    results.append((mod, label, s, d))
                    print(f"  {s:4s} {label}")
                else:
                    # Handle datetime/date returns (VBA Date → COM datetime)
                    if isinstance(vba_r, (datetime, date)):
                        vba_dt = datetime(vba_r.year, vba_r.month, vba_r.day)
                        py_dt = datetime(py_v.year, py_v.month, py_v.day) if isinstance(py_v, (datetime, date)) else None
                        if py_dt is not None:
                            diff = abs((vba_dt - py_dt).total_seconds() / 86400.0)
                            ok = diff < 1.0
                            s = "PASS" if ok else "FAIL"
                            d = "" if ok else f"VBA={vba_dt.date()} PY={py_dt.date()}"
                        else:
                            s, d = "FAIL", f"date vs non-date: VBA={vba_dt.date()} PY={py_v}"
                        results.append((mod, label, s, d))
                        print(f"  {s:4s} {label} VBA={vba_dt.date()} PY={py_dt.date() if py_dt else py_v}")
                    else:
                        vv = _com_to_scalar(vba_r)
                        if vv is None:
                            s, d = "FAIL", "null result"
                        else:
                            ok = abs(vv - float(py_v)) <= tol
                            s = "PASS" if ok else "FAIL"
                            d = "" if ok else f"VBA={vv} PY={py_v}"
                        results.append((mod, label, s, d))
                        print(f"  {s:4s} {label} VBA={vv} PY={py_v}")
            except Exception as exc:
                results.append((mod, label, "FAIL", str(exc)))
                print(f"  FAIL {label} — {exc}")

        passed = sum(1 for r in results if r[2] == "PASS")
        failed = sum(1 for r in results if r[2] == "FAIL")
        print(f"\n{'='*60}")
        print(f"  Manual Examples: {len(results)} total, {passed} PASS, {failed} FAIL")
        print(f"{'='*60}")
        if failed:
            print("  Failures:")
            for m, l, s, d in results:
                if s == "FAIL": print(f"    - {l}: {d}")
        return 0 if failed == 0 else 1
    finally:
        teardown(excel, wb)


if __name__ == "__main__":
    sys.exit(main())

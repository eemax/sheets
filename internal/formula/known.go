package formula

import (
	"sort"
	"strings"
)

// classic lists the functions of the Excel 2007 file format; together with
// the future functions they are every function Excel knows. A call to
// anything else is #NAME? in Excel.
var classic = setOf(`
	ABS ACCRINT ACCRINTM ACOS ACOSH ADDRESS AMORDEGRC AMORLINC AND AREAS ASC ASIN
	ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BAHTTEXT
	BESSELI BESSELJ BESSELK BESSELY BETADIST BETAINV BIN2DEC BIN2HEX BIN2OCT
	BINOMDIST CALL CEILING CELL CHAR CHIDIST CHIINV CHITEST CHOOSE CLEAN CODE
	COLUMN COLUMNS COMBIN COMPLEX CONCATENATE CONFIDENCE CONVERT CORREL COS COSH
	COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD
	COUPNUM COUPPCD COVAR CRITBINOM CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY
	CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE CUMIPMT CUMPRINC DATE DATEDIF
	DATEVALUE DAVERAGE DAY DAYS360 DB DBCS DCOUNT DCOUNTA DDB DEC2BIN DEC2HEX
	DEC2OCT DEGREES DELTA DEVSQ DGET DISC DMAX DMIN DOLLAR DOLLARDE DOLLARFR
	DPRODUCT DSTDEV DSTDEVP DSUM DURATION DVAR DVARP EDATE EFFECT EOMONTH ERF ERFC
	ERROR.TYPE EUROCONVERT EVEN EXACT EXP EXPONDIST FACT FACTDOUBLE FALSE FDIST
	FIND FINDB FINV FISHER FISHERINV FIXED FLOOR FORECAST FREQUENCY FTEST FV
	FVSCHEDULE GAMMADIST GAMMAINV GAMMALN GCD GEOMEAN GESTEP GETPIVOTDATA GROWTH
	HARMEAN HEX2BIN HEX2DEC HEX2OCT HLOOKUP HOUR HYPERLINK HYPGEOMDIST IF IFERROR
	IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMDIV IMEXP IMLN IMLOG10 IMLOG2
	IMPOWER IMPRODUCT IMREAL IMSIN IMSQRT IMSUB IMSUM INDEX INDIRECT INFO INT
	INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISLOGICAL ISNA
	ISNONTEXT ISNUMBER ISODD ISPMT ISREF ISTEXT JIS KURT LARGE LCM LEFT LEFTB LEN
	LENB LINEST LN LOG LOG10 LOGEST LOGINV LOGNORMDIST LOOKUP LOWER MATCH MAX MAXA
	MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINUTE MINVERSE MIRR MMULT MOD MODE
	MONTH MROUND MULTINOMIAL N NA NEGBINOMDIST NETWORKDAYS NOMINAL NORMDIST NORMINV
	NORMSDIST NORMSINV NOT NOW NPER NPV OCT2BIN OCT2DEC OCT2HEX ODD ODDFPRICE
	ODDFYIELD ODDLPRICE ODDLYIELD OFFSET OR PEARSON PERCENTILE PERCENTRANK PERMUT
	PHONETIC PI PMT POISSON POWER PPMT PRICE PRICEDISC PRICEMAT PROB PRODUCT
	PROPER PV QUARTILE QUOTIENT RADIANS RAND RANDBETWEEN RANK RATE RECEIVED
	REGISTER.ID REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP
	ROW ROWS RSQ RTD SEARCH SEARCHB SECOND SERIESSUM SIGN SIN SINH SKEW SLN SLOPE
	SMALL SQL.REQUEST SQRT SQRTPI STANDARDIZE STDEV STDEVA STDEVP STDEVPA STEYX
	SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2
	SUMXMY2 SYD T TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TIME TIMEVALUE
	TINV TODAY TRANSPOSE TREND TRIM TRIMMEAN TRUE TRUNC TTEST TYPE UPPER VALUE VAR
	VARA VARP VARPA VDB VLOOKUP WEEKDAY WEEKNUM WEIBULL WORKDAY XIRR XNPV YEAR
	YEARFRAC YIELD YIELDDISC YIELDMAT ZTEST`)

// IsFunction reports whether Excel knows a function by this name.
func IsFunction(name string) bool {
	n := strings.ToUpper(bare(name))
	_, fut := future[n]
	return classic[n] || fut
}

// UnknownFunctions lists called names that are neither Excel functions,
// LET/LAMBDA names, nor accepted by isName (workbook names, which can hold
// LAMBDAs). Excel shows #NAME? for these.
func UnknownFunctions(f string, isName func(string) bool) []string {
	ts := lex(f)
	params := lambdaParams(ts)
	var out []string
	seen := map[string]bool{}
	for i, t := range ts {
		if !t.call || i > 0 && ts[i-1].text == "!" {
			continue
		}
		n := strings.ToUpper(bare(t.text))
		if IsFunction(n) || params[n] || seen[n] || isName != nil && isName(bare(t.text)) {
			continue
		}
		seen[n] = true
		out = append(out, bare(t.text))
	}
	return out
}

var lambdaFamily = setOf(`LET LAMBDA MAP REDUCE SCAN BYROW BYCOL MAKEARRAY ISOMITTED`)

// LambdaFamily returns the first LET/LAMBDA-style function a formula uses,
// or "". The sheets calculator evaluates none of them.
func LambdaFamily(f string) string {
	for _, t := range lex(f) {
		if n := strings.ToUpper(bare(t.text)); t.call && lambdaFamily[n] {
			return n
		}
	}
	return ""
}

// Functions lists every function Excel knows, by its plain name.
func Functions() []string {
	out := make([]string, 0, len(classic)+len(future))
	for f := range classic {
		out = append(out, f)
	}
	for f := range future {
		out = append(out, f)
	}
	sort.Strings(out)
	return out
}

// CalledFunctions lists the functions a formula calls, by plain name, upper
// case; names that are not Excel functions (a LAMBDA kept in a workbook
// name) included.
func CalledFunctions(f string) []string {
	var out []string
	seen := map[string]bool{}
	ts := lex(f)
	for i, t := range ts {
		if !t.call || i > 0 && ts[i-1].text == "!" {
			continue
		}
		if n := strings.ToUpper(bare(t.text)); !seen[n] {
			seen[n] = true
			out = append(out, n)
		}
	}
	return out
}

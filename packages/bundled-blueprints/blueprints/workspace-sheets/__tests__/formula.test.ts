// @vitest-environment node
import { describe, expect, it } from "vitest";
import { type Ast, CellError, parseFormula, serializeAst, tokenize, unquoteSheetName, unwrapParens } from "../files/lib/formula.ts";

describe("the formula tokenizer", () => {
  it("reads numbers, strings, operators, references and quoted sheet names", () => {
    expect(tokenize('SUM(A1:B2, 1.5e3) & "a""b" <> \'My Sheet\'!C3')).toEqual([
      { t: "word", v: "SUM" }, { t: "lp" }, { t: "word", v: "A1" }, { t: "colon" }, { t: "word", v: "B2" },
      { t: "comma" }, { t: "num", v: 1500 }, { t: "rp" }, { t: "op", v: "&" }, { t: "str", v: 'a"b' },
      { t: "op", v: "<>" }, { t: "word", v: "'My Sheet'!C3" },
    ]);
  });

  it("reads a single-quoted run as text unless a sheet reference follows it", () => {
    expect(tokenize("'abc'")).toEqual([{ t: "str", v: "abc" }]);
    expect(tokenize("'it''s' & 'Q1'!A1")).toEqual([
      { t: "str", v: "it's" }, { t: "op", v: "&" }, { t: "word", v: "'Q1'!A1" },
    ]);
    expect(tokenize("''")).toEqual([{ t: "str", v: "" }]);
  });

  it("reads a backslash as an ordinary character, as Excel does", () => {
    expect(tokenize('"C:\\"&A1')).toEqual([{ t: "str", v: "C:\\" }, { t: "op", v: "&" }, { t: "word", v: "A1" }]);
    expect(serializeAst(parseFormula('SUBSTITUTE(A1,"\\","/")'))).toBe('SUBSTITUTE(A1,"\\","/")');
    expect(tokenize("'Data\\'!A1")).toEqual([{ t: "word", v: "'Data\\'!A1" }]);
  });

  it("keeps a doubled quote in a sheet name escaped, so the word writes back unchanged", () => {
    expect(tokenize("'O''Brien'!B2")).toEqual([{ t: "word", v: "'O''Brien'!B2" }]);
    expect(serializeAst(parseFormula("'O''Brien'!B2+1"))).toBe("'O''Brien'!B2+1");
    expect(unquoteSheetName("'O''Brien'")).toBe("O'Brien");
    expect(unquoteSheetName("Plain")).toBe("Plain");
  });
});

describe("the formula parser", () => {
  it("builds precedence-aware trees and writes them back", () => {
    const ast = parseFormula("1+2*A1^2%");
    expect(ast).toEqual({
      k: "bin", op: "+", a: { k: "num", v: 1 },
      b: { k: "bin", op: "*", a: { k: "num", v: 2 }, b: { k: "bin", op: "^", a: { k: "ref", ref: "A1" }, b: { k: "pct", a: { k: "num", v: 2 } } } },
    });
    expect(serializeAst(ast)).toBe("1+2*A1^2%");
    expect(serializeAst(parseFormula('IF(A1>=3, "yes", TRUE)'))).toBe('IF(A1>=3,"yes",TRUE)');
  });

  it("parses a range as two references", () => {
    expect(parseFormula("SUM(A1:B2)")).toEqual({
      k: "call", name: "SUM", args: [{ k: "range", a: "A1", b: "B2" }],
    });
    expect(serializeAst(parseFormula("Sheet2!A1:'Other Sheet'!B9")))
      .toBe("Sheet2!A1:'Other Sheet'!B9");
  });

  it.each(["A1:5", "A1:(", "A1:", "SUM(A1:)"])("rejects %s, whose range has no end reference", (src) => {
    let thrown: unknown;
    try { parseFormula(src); } catch (e) { thrown = e; }
    expect(thrown).toBeInstanceOf(CellError);
    expect(String(thrown)).toBe("#VALUE!");
  });

  it("rejects an empty or dangling expression as #VALUE!", () => {
    for (const src of ["", "1+", "*2"]) {
      let thrown: unknown;
      try { parseFormula(src); } catch (e) { thrown = e; }
      expect(thrown, src).toBeInstanceOf(CellError);
    }
  });

  it.each(["1 2", "A1)", "SUM(A1))", "A1 B1"])("rejects %s, which has tokens after a complete expression", (src) => {
    let thrown: unknown;
    try { parseFormula(src); } catch (e) { thrown = e; }
    expect(thrown).toBeInstanceOf(CellError);
    expect(String(thrown)).toBe("#VALUE!");
  });

  it("still accepts a call whose closing parenthesis is missing", () => {
    expect(parseFormula("SUM(A1")).toEqual({ k: "call", name: "SUM", args: [{ k: "ref", ref: "A1" }] });
  });

  it("keeps parentheses when writing a formula back", () => {
    // Rows and columns are inserted by rewriting each formula's references through the
    // serializer, so a dropped pair of parentheses would silently change what a cell computes.
    expect(serializeAst(parseFormula("(A1+A2)*3"))).toBe("(A1+A2)*3");
    expect(serializeAst(parseFormula("-(A1)"))).toBe("-(A1)");
    expect(parseFormula("(1+2)*3")).toEqual({
      k: "bin", op: "*", a: { k: "paren", a: { k: "bin", op: "+", a: { k: "num", v: 1 }, b: { k: "num", v: 2 } } }, b: { k: "num", v: 3 },
    });
  });

  it("sees through grouping parentheses on request", () => {
    // ROW and COLUMN read their argument's shape rather than its value, so `ROW((A5))` has to find
    // the reference behind the grouping the parser keeps for the serializer.
    expect(unwrapParens(parseFormula("((A5))"))).toEqual({ k: "ref", ref: "A5" });
    expect(unwrapParens(parseFormula("(A1:B2)"))).toEqual({ k: "range", a: "A1", b: "B2" });
    const ref: Ast = { k: "ref", ref: "A5" };
    expect(unwrapParens(ref)).toBe(ref);
  });

  it("round-trips every node kind through the serializer", () => {
    const sources = ["-A1", "A1%", "(1+2)*3", '"quoted ""text"""', "FALSE", "F(A1,B2:C3)"];
    for (const src of sources) {
      const ast: Ast = parseFormula(src);
      expect(parseFormula(serializeAst(ast)), src).toEqual(ast);
    }
  });
});

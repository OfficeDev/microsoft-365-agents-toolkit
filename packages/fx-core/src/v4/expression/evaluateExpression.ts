// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";

/**
 * The v4 closed-grammar expression evaluator — the single shared evaluator.
 *
 * Spec: docs/03-specs/operations/scaffolding/evaluate-expression.md
 * Decision: docs/02-architecture/adr/ADR-0016-declarative-template-format.md (decision 7)
 *
 * This module is part of the v4 world. It imports no v3 symbol. The evaluator
 * is pure and synchronous by design (spec INV-4) — no fs / http / clock — which
 * is exactly what keeps it `InMemoryRuntime`-fuzzable.
 */

const SOURCE = "Scaffold";

/** A value the evaluator yields: a boolean (a `when` / `condition`) or a string (a `{expr}` value). */
export type EvalValue = boolean | string;

/**
 * The grammar's `null` literal and the value of a declared-but-unanswered
 * identifier (spec INPUT-12 of `collect-inputs`: `mcpServerUrl == null`).
 * Distinct from the empty string — a unique symbol so `===` is an identity test
 * no answered string value can satisfy. `collect-inputs` seeds an unanswered but
 * declared option id with this marker so `x == null` is a presence test, while a
 * truly undeclared identifier (a typo) is still an `EXPR_UNDECLARED_IDENTIFIER`.
 */
export const NULL_VALUE: unique symbol = Symbol("v4.expr.null");

/** A resolved identifier's value: a string, or `NULL_VALUE` (unanswered / absent). */
export type ScopeValue = string | typeof NULL_VALUE;

/**
 * The resolved identifier map (spec INV-3): `optionsSchema.properties` ∪
 * step-produced fields ∪ the caller-injected floor. An identifier absent from
 * this map is a `SystemError`, never a silent falsy default. A declared id whose
 * answer is not yet known maps to `NULL_VALUE` (vs. truly absent = undeclared).
 */
export type Scope = Record<string, ScopeValue>;

/**
 * A whitelisted, side-effect-free function
 * (`safeUpper` / `safeLower` / `safeServer` / `safeAlphanumeric` / `surface` /
 * `mcpNamespace` / `mcpAuthRef` — ADR-0016 decision 3). Every entry is total
 * and pure; the whitelist is engine-owned and cannot be extended by an author.
 */
export type WhitelistFn = (...args: string[]) => string;

/**
 * The narrowest port in the domain: the fixed pure-function whitelist plus the
 * read-only feature-flag reader behind `featureFlag('…')`. Deliberately carries
 * no `fs` / `http` / `clock` face (spec INV-4).
 */
export interface ExpressionRuntimePort {
  /** The fixed function whitelist; `undefined` for any non-whitelisted name. */
  functions(name: string): WhitelistFn | undefined;
  /** The feature-flag reader behind `featureFlag('…')` (env-backed, read-only). */
  flags(name: string): boolean;
}

/**
 * An authored closed form: the raw `expr` grammar, or one of the sugar forms
 * that desugars losslessly into it at load (spec INV-1, INV-5).
 */
export type ExpressionNode =
  | { expr: string }
  | { equals: Record<string, string> }
  | { enum: Record<string, string[]> }
  | { anyOf: ExpressionNode[] }
  | { featureFlag: string }
  | { capability: string }
  | { from: string };

/**
 * `SystemError` names. Reaching any of these at runtime is an engine bug — the
 * build-time typed-context check (ADR-0016 decision 9) and
 * `validate-template-package` should have caught it (spec INV-2, INV-3).
 */
export const EXPR_UNDECLARED_IDENTIFIER = "ExprUndeclaredIdentifier";
export const EXPR_NON_WHITELISTED_FUNCTION = "ExprNonWhitelistedFunction";
export const EXPR_PARSE_ERROR = "ExprParseError";

/**
 * Evaluate one authored closed-form expression against a resolved name space.
 *
 * @param node  the authored `when` / `condition` / `{expr}` (raw or sugar form)
 * @param scope the resolved identifier map (spec INV-3)
 * @param port  the pure function whitelist + feature-flag reader (spec INV-4)
 * @returns `ok(boolean | string)`, or a `SystemError` for an undeclared
 *          identifier, a non-whitelisted call, or a parse failure.
 */
export function evaluateExpression(
  node: ExpressionNode,
  scope: Scope,
  port: ExpressionRuntimePort
): Result<EvalValue, FxError> {
  // Internal parse / scope / whitelist failures are signalled with a private
  // `EvalError` and converted to a `SystemError` here, at the single public
  // boundary — no throw escapes this Result contract.
  try {
    const expr = desugarToExpr(node);
    const ast = parse(tokenize(expr));
    const value = evalAst(ast, scope, port);
    // A top-level `null` (an unanswered declared id in value context) renders as
    // the empty string; a condition never yields `null` at top level (a
    // comparison yields a boolean), so this only affects a `{from}` / `{expr}`
    // value reference to an unanswered id.
    return ok(value === NULL_VALUE ? "" : value);
  } catch (e) {
    const name = e instanceof EvalError ? e.code : EXPR_PARSE_ERROR;
    const message = e instanceof Error ? e.message : String(e);
    return err(new SystemError({ source: SOURCE, name, message }));
  }
}

// --- Internal: one grammar, one tokenizer, one evaluator (spec INV-1). ---

/** A typed, internal-only control-flow signal converted to a `SystemError` at the public boundary. */
class EvalError extends Error {
  public readonly code: string;
  constructor(code: string, message: string) {
    super(message);
    this.code = code;
  }
}

/** Lower each sugar form to the single `expr` grammar (spec INV-5, loss-less). */
function desugarToExpr(node: ExpressionNode): string {
  if ("expr" in node) {
    return node.expr;
  }
  if ("from" in node) {
    return node.from; // value context: a verbatim identifier copy
  }
  if ("featureFlag" in node) {
    return `featureFlag('${node.featureFlag}')`;
  }
  if ("capability" in node) {
    return `capability == '${node.capability}'`;
  }
  if ("equals" in node) {
    const [key, value] = Object.entries(node.equals)[0];
    return `${key} == '${value}'`;
  }
  if ("enum" in node) {
    const [key, values] = Object.entries(node.enum)[0];
    return values.map((v) => `${key} == '${v}'`).join(" || ");
  }
  // anyOf
  return node.anyOf.map((c) => `(${desugarToExpr(c)})`).join(" || ");
}

type Token =
  | { t: "lparen" }
  | { t: "rparen" }
  | { t: "comma" }
  | { t: "eq" }
  | { t: "ne" }
  | { t: "not" }
  | { t: "and" }
  | { t: "or" }
  | { t: "str"; v: string }
  | { t: "ident"; v: string };

function tokenize(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      i++;
    } else if (c === "(") {
      tokens.push({ t: "lparen" });
      i++;
    } else if (c === ")") {
      tokens.push({ t: "rparen" });
      i++;
    } else if (c === ",") {
      tokens.push({ t: "comma" });
      i++;
    } else if (c === "=" && src[i + 1] === "=") {
      tokens.push({ t: "eq" });
      i += 2;
    } else if (c === "!" && src[i + 1] === "=") {
      tokens.push({ t: "ne" });
      i += 2;
    } else if (c === "!") {
      tokens.push({ t: "not" });
      i++;
    } else if (c === "&" && src[i + 1] === "&") {
      tokens.push({ t: "and" });
      i += 2;
    } else if (c === "|" && src[i + 1] === "|") {
      tokens.push({ t: "or" });
      i += 2;
    } else if (c === "'") {
      let j = i + 1;
      let s = "";
      while (j < src.length && src[j] !== "'") {
        s += src[j];
        j++;
      }
      if (j >= src.length) {
        throw new EvalError(EXPR_PARSE_ERROR, "unterminated string literal");
      }
      tokens.push({ t: "str", v: s });
      i = j + 1;
    } else if (/[A-Za-z_]/.test(c)) {
      let j = i;
      let s = "";
      while (j < src.length && /[A-Za-z0-9_]/.test(src[j])) {
        s += src[j];
        j++;
      }
      tokens.push({ t: "ident", v: s });
      i = j;
    } else {
      throw new EvalError(EXPR_PARSE_ERROR, `unexpected character '${c}'`);
    }
  }
  return tokens;
}

type Ast =
  | { k: "str"; v: string }
  | { k: "null" }
  | { k: "ident"; v: string }
  | { k: "call"; name: string; args: Ast[] }
  | { k: "eq"; left: Ast; right: Ast; neg: boolean }
  | { k: "not"; operand: Ast }
  | { k: "and"; left: Ast; right: Ast }
  | { k: "or"; left: Ast; right: Ast };

/** Recursive-descent parse: or → and → equality → unary → primary (call | group | literal | identifier). */
function parse(tokens: Token[]): Ast {
  let pos = 0;
  const peek = (): Token | undefined => tokens[pos];
  const eat = (): Token => {
    const t = tokens[pos];
    if (!t) {
      throw new EvalError(EXPR_PARSE_ERROR, "unexpected end of expression");
    }
    pos++;
    return t;
  };

  function parseOr(): Ast {
    let left = parseAnd();
    while (peek()?.t === "or") {
      eat();
      left = { k: "or", left, right: parseAnd() };
    }
    return left;
  }
  function parseAnd(): Ast {
    let left = parseEquality();
    while (peek()?.t === "and") {
      eat();
      left = { k: "and", left, right: parseEquality() };
    }
    return left;
  }
  function parseEquality(): Ast {
    let left = parseUnary();
    let op = peek();
    while (op?.t === "eq" || op?.t === "ne") {
      eat();
      left = { k: "eq", left, right: parseUnary(), neg: op.t === "ne" };
      op = peek();
    }
    return left;
  }
  function parseUnary(): Ast {
    if (peek()?.t === "not") {
      eat();
      return { k: "not", operand: parseUnary() };
    }
    return parsePrimary();
  }
  function parsePrimary(): Ast {
    const t = peek();
    if (!t) {
      throw new EvalError(EXPR_PARSE_ERROR, "unexpected end of expression");
    }
    if (t.t === "lparen") {
      eat();
      const inner = parseOr();
      if (peek()?.t !== "rparen") {
        throw new EvalError(EXPR_PARSE_ERROR, "missing closing ')'");
      }
      eat();
      return inner;
    }
    if (t.t === "str") {
      eat();
      return { k: "str", v: t.v };
    }
    if (t.t === "ident") {
      eat();
      if (t.v === "null") {
        // `null` is a reserved literal (the absent value), never an identifier
        // or a callable name.
        return { k: "null" };
      }
      if (peek()?.t === "lparen") {
        eat();
        const args: Ast[] = [];
        if (peek()?.t !== "rparen") {
          args.push(parseOr());
          while (peek()?.t === "comma") {
            eat();
            args.push(parseOr());
          }
        }
        if (peek()?.t !== "rparen") {
          throw new EvalError(EXPR_PARSE_ERROR, "missing closing ')' in call");
        }
        eat();
        return { k: "call", name: t.v, args };
      }
      return { k: "ident", v: t.v };
    }
    throw new EvalError(EXPR_PARSE_ERROR, "unexpected token");
  }

  const ast = parseOr();
  if (pos !== tokens.length) {
    throw new EvalError(EXPR_PARSE_ERROR, "trailing tokens after expression");
  }
  return ast;
}

/** The internal value domain: a public `EvalValue` plus the `null` marker. */
type Val = EvalValue | typeof NULL_VALUE;

function asString(v: Val): string {
  if (v === NULL_VALUE) {
    return "";
  }
  return typeof v === "string" ? v : String(v);
}
function asBoolean(v: Val): boolean {
  if (typeof v === "boolean") {
    return v;
  }
  throw new EvalError(EXPR_PARSE_ERROR, "expected a boolean operand");
}

/** Evaluate the AST against `scope` + `port`; `&&` / `||` short-circuit so a skipped `featureFlag` is never read. */
function evalAst(ast: Ast, scope: Scope, port: ExpressionRuntimePort): Val {
  switch (ast.k) {
    case "str":
      return ast.v;
    case "null":
      return NULL_VALUE;
    case "ident": {
      if (!(ast.v in scope)) {
        throw new EvalError(EXPR_UNDECLARED_IDENTIFIER, `identifier '${ast.v}' is not in scope`);
      }
      return scope[ast.v];
    }
    case "call": {
      if (ast.name === "featureFlag") {
        return port.flags(asString(evalAst(ast.args[0], scope, port)));
      }
      const fn = port.functions(ast.name);
      if (!fn) {
        throw new EvalError(
          EXPR_NON_WHITELISTED_FUNCTION,
          `function '${ast.name}' is not in the whitelist`
        );
      }
      return fn(...ast.args.map((a) => asString(evalAst(a, scope, port))));
    }
    case "eq": {
      const equal = evalAst(ast.left, scope, port) === evalAst(ast.right, scope, port);
      return ast.neg ? !equal : equal;
    }
    case "and":
      return (
        asBoolean(evalAst(ast.left, scope, port)) && asBoolean(evalAst(ast.right, scope, port))
      );
    case "or":
      return (
        asBoolean(evalAst(ast.left, scope, port)) || asBoolean(evalAst(ast.right, scope, port))
      );
    case "not":
      return !asBoolean(evalAst(ast.operand, scope, port));
    default:
      throw new EvalError(EXPR_PARSE_ERROR, "unknown expression node");
  }
}

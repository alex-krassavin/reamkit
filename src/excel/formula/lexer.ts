// Formula lexer (E-SHEET W9) — turns a formula string into a flat token stream
// for the Pratt parser. Recognises numbers, quoted strings, error literals
// (#REF! …), "words" (cell references, function names, defined names, TRUE/
// FALSE — classified later by the parser), and the operator/punctuation set.
// Whitespace separates tokens but is otherwise dropped (the space intersection
// operator is not supported — irrelevant for conditional-format expressions).

export type TokenKind =
  | 'num'
  | 'str'
  | 'err'
  | 'word' // a cell ref, function name, defined name, or TRUE/FALSE — parser decides
  | 'sheetq' // a 'quoted sheet name' (only valid immediately before a `!` qualifier)
  | 'op' // + - * / ^ & = <> < > <= >= % : ! ( ) ,
  | 'eof';

export interface Token {
  readonly kind: TokenKind;
  readonly text: string;
}

// Conditional-format formulas are short; cap the source so a crafted workbook
// cannot make us tokenize an enormous string. Anything longer simply fails to
// parse and the rule becomes a no-op (never a hang, never a misrender).
const MAX_SOURCE = 8192;

const KNOWN_ERRORS: ReadonlySet<string> = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
]);

export class LexError extends Error {}

function isDigit(ch: string): boolean {
  return ch >= '0' && ch <= '9';
}

function isWordStart(ch: string): boolean {
  return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || ch === '_' || ch === '$';
}

function isWordPart(ch: string): boolean {
  return isWordStart(ch) || isDigit(ch) || ch === '.';
}

export function tokenize(src: string): Array<Token> {
  if (src.length > MAX_SOURCE) throw new LexError('formula too long');
  const out: Array<Token> = [];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const ch = src[i]!;
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      i++;
      continue;
    }
    // A 'single-quoted sheet name' — '' is an escaped quote. Only meaningful right
    // before a `!` (Sheet2!A1 / 'My Sheet'!A1); the parser enforces that.
    if (ch === "'") {
      let s = '';
      i++;
      while (i < n) {
        if (src[i] === "'") {
          if (src[i + 1] === "'") {
            s += "'";
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += src[i];
        i++;
      }
      out.push({ kind: 'sheetq', text: s });
      continue;
    }
    // Quoted string — "" is an escaped quote inside the literal.
    if (ch === '"') {
      let s = '';
      i++;
      while (i < n) {
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            s += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += src[i];
        i++;
      }
      out.push({ kind: 'str', text: s });
      continue;
    }
    // Error literal — read the maximal run of error characters after '#'.
    if (ch === '#') {
      let j = i + 1;
      while (j < n && /[A-Za-z0-9/!?]/.test(src[j]!)) j++;
      const lit = src.slice(i, j).toUpperCase();
      if (!KNOWN_ERRORS.has(lit)) throw new LexError(`bad error literal ${lit}`);
      out.push({ kind: 'err', text: lit });
      i = j;
      continue;
    }
    // Number — digits with an optional fraction and exponent (or a leading dot).
    if (isDigit(ch) || (ch === '.' && i + 1 < n && isDigit(src[i + 1]!))) {
      let j = i;
      while (j < n && isDigit(src[j]!)) j++;
      if (src[j] === '.') {
        j++;
        while (j < n && isDigit(src[j]!)) j++;
      }
      if (src[j] === 'e' || src[j] === 'E') {
        let k = j + 1;
        if (src[k] === '+' || src[k] === '-') k++;
        if (k < n && isDigit(src[k]!)) {
          k++;
          while (k < n && isDigit(src[k]!)) k++;
          j = k;
        }
      }
      out.push({ kind: 'num', text: src.slice(i, j) });
      i = j;
      continue;
    }
    // Word — cell ref / function name / defined name / TRUE / FALSE.
    if (isWordStart(ch)) {
      let j = i + 1;
      while (j < n && isWordPart(src[j]!)) j++;
      out.push({ kind: 'word', text: src.slice(i, j) });
      i = j;
      continue;
    }
    // Two-character comparison operators.
    if (ch === '<' && src[i + 1] === '=') {
      out.push({ kind: 'op', text: '<=' });
      i += 2;
      continue;
    }
    if (ch === '>' && src[i + 1] === '=') {
      out.push({ kind: 'op', text: '>=' });
      i += 2;
      continue;
    }
    if (ch === '<' && src[i + 1] === '>') {
      out.push({ kind: 'op', text: '<>' });
      i += 2;
      continue;
    }
    // Single-character operators / punctuation (`!` separates a sheet qualifier).
    if ('+-*/^&=<>%:(),!'.includes(ch)) {
      out.push({ kind: 'op', text: ch });
      i++;
      continue;
    }
    throw new LexError(`unexpected character ${JSON.stringify(ch)}`);
  }
  out.push({ kind: 'eof', text: '' });
  return out;
}

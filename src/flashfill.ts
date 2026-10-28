// Flash Fill, as Excel's Ctrl+E: from a few results typed next to the data, works out how each
// is made from the cells beside it (which words, in what case, with what between them) and
// fills in the rest of the column the same way. "Kari Nordmann" → "K. Nordmann", "kari@uio.no" →
// "Kari", "2024-03-15" → "15.03".

/** One piece of a result: a word of a cell (or its first letter), in some case, or fixed text. */
type Piece =
  | { kind: "text"; text: string }
  | { kind: "word"; col: number; word: number; fromEnd: boolean; initial: boolean; case: "same" | "upper" | "lower" | "proper" };

export type Program = Piece[];

/** The words (letters and digits) of a cell, in order. */
const wordsOf = (s: string) => s.match(/[\p{L}\p{N}]+/gu) ?? [];

const CASES = ["same", "proper", "upper", "lower"] as const;
const applyCase = (s: string, c: (typeof CASES)[number]) =>
  c === "upper" ? s.toUpperCase() : c === "lower" ? s.toLowerCase() : c === "proper" ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : s;

/** What a word piece gives for a row's cells, or null if that word isn't there. */
function wordOf(piece: Extract<Piece, { kind: "word" }>, inputs: string[]): string | null {
  const words = wordsOf(inputs[piece.col] ?? "");
  const w = piece.fromEnd ? words[words.length - 1 - piece.word] : words[piece.word];
  if (w === undefined) return null;
  return applyCase(piece.initial ? w.charAt(0) : w, piece.case);
}

/** A result for a row's cells, or null if the program doesn't fit them. */
export function run(program: Program, inputs: string[]): string | null {
  let out = "";
  for (const piece of program) {
    if (piece.kind === "text") out += piece.text;
    else {
      const w = wordOf(piece, inputs);
      if (w === null) return null;
      out += w;
    }
  }
  return out;
}

/** Every way the word pieces could make up the start of `rest`, for one example's cells. */
function wordPieces(inputs: string[], rest: string): { piece: Piece; length: number }[] {
  const found: { piece: Piece; length: number }[] = [];
  inputs.forEach((cell, col) => {
    const words = wordsOf(cell);
    words.forEach((w, i) => {
      for (const initial of [false, true]) {
        for (const c of CASES) {
          const text = applyCase(initial ? w.charAt(0) : w, c);
          if (!text || !rest.startsWith(text)) continue;
          // Counted from the start and from the end (other rows decide); with one example, the
          // last word is taken as "the last" (a surname), any other as counted from the start.
          const fromStart = { piece: { kind: "word", col, word: i, fromEnd: false, initial, case: c } as Piece, length: text.length };
          const fromEnd = { piece: { kind: "word", col, word: words.length - 1 - i, fromEnd: true, initial, case: c } as Piece, length: text.length };
          found.push(...(i === words.length - 1 && i > 0 ? [fromEnd, fromStart] : [fromStart, fromEnd]));
        }
      }
    });
  });
  // Whole words before initials, so "Nordmann" is read as a word, not as "N" and "ordmann".
  return found.sort((a, b) => b.length - a.length);
}

/**
 * The simplest program that turns each example's cells into its result, or null if none is found.
 * Programs are tried from the first example, shortest first, and kept if they make every other.
 */
export function learn(examples: { inputs: string[]; output: string }[]): Program | null {
  if (!examples.length) return null;
  const [first, ...others] = examples;
  const fits = (p: Program) => others.every((e) => run(p, e.inputs) === e.output);
  let tried = 0;
  let best: Program | null = null;
  let bestCost = Infinity;
  // As little fixed text as possible (it's what the cells don't explain), then as few pieces.
  const cost = (program: Program) => program.reduce((sum, p) => sum + (p.kind === "text" ? p.text.length * 10 : 0) + 1, 0);
  const search = (at: number, program: Program) => {
    if (tried++ > 20000 || cost(program) >= bestCost) return;
    if (at === first.output.length) {
      if (run(program, first.inputs) === first.output && fits(program)) [best, bestCost] = [[...program], cost(program)];
      return;
    }
    const rest = first.output.slice(at);
    for (const { piece, length } of wordPieces(first.inputs, rest)) {
      program.push(piece);
      search(at + length, program);
      program.pop();
    }
    // Or the next character is fixed text (joined to fixed text before it).
    const last = program[program.length - 1];
    if (last?.kind === "text") {
      program[program.length - 1] = { kind: "text", text: last.text + rest[0] };
      search(at + 1, program);
      program[program.length - 1] = last;
    } else {
      program.push({ kind: "text", text: rest[0] });
      search(at + 1, program);
      program.pop();
    }
  };
  search(0, []);
  // Only fixed text: that's a copy, not a pattern.
  return best && (best as Program).some((p) => p.kind === "word") ? best : null;
}

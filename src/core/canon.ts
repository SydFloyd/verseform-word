import { ntVerseCounts, otVerseCounts } from "bible-tools";

export type CanonBook = {
  id: string;
  name: string;
  aliases: readonly string[];
  verseCounts: readonly number[];
  unavailableVerses?: Readonly<Record<number, readonly number[]>>;
};

export type CanonMetadata = {
  translationId: string;
  version: number;
  books: readonly CanonBook[];
};

type BookDefinition = readonly [id: string, name: string, aliases: readonly string[]];

function numbered(number: 1 | 2 | 3, name: string, short: readonly string[]): string[] {
  const roman = number === 1 ? "I" : number === 2 ? "II" : "III";
  const word = number === 1 ? "First" : number === 2 ? "Second" : "Third";
  return [...new Set([
    `${number} ${name}`, `${number}${name}`, `${roman} ${name}`, `${word} ${name}`,
    ...short.flatMap((alias) => [
      `${number} ${alias}`, `${number}${alias}`, `${roman} ${alias}`, `${word} ${alias}`,
    ]),
  ])];
}

const definitions: readonly BookDefinition[] = [
  ["GEN", "Genesis", ["Gen", "Gn"]],
  ["EXO", "Exodus", ["Exod", "Exo"]],
  ["LEV", "Leviticus", ["Lev", "Lv"]],
  ["NUM", "Numbers", ["Num", "Nm"]],
  ["DEU", "Deuteronomy", ["Deut", "Deu", "Dt"]],
  ["JOS", "Joshua", ["Josh", "Jos"]],
  ["JDG", "Judges", ["Judg", "Jdg"]],
  ["RUT", "Ruth", ["Rth"]],
  ["1SA", "1 Samuel", numbered(1, "Samuel", ["Sam", "Sa"])],
  ["2SA", "2 Samuel", numbered(2, "Samuel", ["Sam", "Sa"])],
  ["1KI", "1 Kings", numbered(1, "Kings", ["Kgs", "Ki"])],
  ["2KI", "2 Kings", numbered(2, "Kings", ["Kgs", "Ki"])],
  ["1CH", "1 Chronicles", numbered(1, "Chronicles", ["Chron", "Chr"] )],
  ["2CH", "2 Chronicles", numbered(2, "Chronicles", ["Chron", "Chr"] )],
  ["EZR", "Ezra", ["Ezr"]],
  ["NEH", "Nehemiah", ["Neh"]],
  ["EST", "Esther", ["Esth", "Est"]],
  ["JOB", "Job", []],
  ["PSA", "Psalms", ["Psalm", "Psa", "Ps"]],
  ["PRO", "Proverbs", ["Prov", "Pro"]],
  ["ECC", "Ecclesiastes", ["Eccl", "Eccles", "Ecc"]],
  ["SNG", "Song of Solomon", ["Song of Songs", "Song", "Canticles", "SOS"]],
  ["ISA", "Isaiah", ["Isa"]],
  ["JER", "Jeremiah", ["Jer"]],
  ["LAM", "Lamentations", ["Lam"]],
  ["EZK", "Ezekiel", ["Ezek", "Eze"]],
  ["DAN", "Daniel", ["Dan"]],
  ["HOS", "Hosea", ["Hos"]],
  ["JOL", "Joel", ["Joe"]],
  ["AMO", "Amos", ["Amo"]],
  ["OBA", "Obadiah", ["Obad", "Oba"]],
  ["JON", "Jonah", ["Jon"]],
  ["MIC", "Micah", ["Mic"]],
  ["NAM", "Nahum", ["Nah"]],
  ["HAB", "Habakkuk", ["Hab"]],
  ["ZEP", "Zephaniah", ["Zeph", "Zep"]],
  ["HAG", "Haggai", ["Hag"]],
  ["ZEC", "Zechariah", ["Zech", "Zec"]],
  ["MAL", "Malachi", ["Mal"]],
  ["MAT", "Matthew", ["Matt", "Mat"]],
  ["MRK", "Mark", ["Mrk", "Mk"]],
  ["LUK", "Luke", ["Luk", "Lk"]],
  ["JHN", "John", ["Jhn", "Jn"]],
  ["ACT", "Acts", ["Act"]],
  ["ROM", "Romans", ["Rom"]],
  ["1CO", "1 Corinthians", numbered(1, "Corinthians", ["Cor", "Corin"])],
  ["2CO", "2 Corinthians", numbered(2, "Corinthians", ["Cor", "Corin"])],
  ["GAL", "Galatians", ["Gal"]],
  ["EPH", "Ephesians", ["Eph"]],
  ["PHP", "Philippians", ["Phil", "Php"]],
  ["COL", "Colossians", ["Col"]],
  ["1TH", "1 Thessalonians", numbered(1, "Thessalonians", ["Thess", "Thes"])],
  ["2TH", "2 Thessalonians", numbered(2, "Thessalonians", ["Thess", "Thes"])],
  ["1TI", "1 Timothy", numbered(1, "Timothy", ["Tim"] )],
  ["2TI", "2 Timothy", numbered(2, "Timothy", ["Tim"] )],
  ["TIT", "Titus", ["Tit"]],
  ["PHM", "Philemon", ["Philem", "Phlm", "Phm"]],
  ["HEB", "Hebrews", ["Heb"]],
  ["JAS", "James", ["Jas", "Jam"]],
  ["1PE", "1 Peter", numbered(1, "Peter", ["Pet"] )],
  ["2PE", "2 Peter", numbered(2, "Peter", ["Pet"] )],
  ["1JN", "1 John", numbered(1, "John", ["Jn", "Jhn"] )],
  ["2JN", "2 John", numbered(2, "John", ["Jn", "Jhn"] )],
  ["3JN", "3 John", numbered(3, "John", ["Jn", "Jhn"] )],
  ["JUD", "Jude", ["Jd"]],
  ["REV", "Revelation", ["Rev"]],
];

const standardCounts = [...otVerseCounts, ...ntVerseCounts].map((counts) => [...counts]);

if (definitions.length !== 66 || standardCounts.length !== definitions.length) {
  throw new Error("The bundled canon metadata is incomplete.");
}

function booksForCounts(): CanonBook[] {
  return definitions.map(([id, name, aliases], index) => ({
    id,
    name,
    aliases: [...new Set(aliases.filter((alias) => alias !== name))],
    verseCounts: [...standardCounts[index]!],
  }));
}

export const STANDARD_CANON: CanonMetadata = {
  translationId: "STANDARD-66",
  version: 1,
  books: booksForCounts(),
};

export function canonBook(canon: CanonMetadata, bookId: string): CanonBook | undefined {
  return canon.books.find((book) => book.id === bookId);
}

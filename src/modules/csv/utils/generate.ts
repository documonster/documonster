/**
 * CSV Data Generator
 *
 * A lightweight utility for generating random CSV data.
 * Useful for testing, benchmarking, and creating mock datasets.
 *
 * Features:
 * - Seeded pseudo-random number generation for reproducible results
 * - Built-in column types: string, int, float, bool, date, uuid, email, etc.
 * - Custom generator function support
 * - Sync and async iteration support
 * - Cross-platform (Node.js and Browser)
 */

import { CsvArgumentError } from "@csv/errors";
import { uuidV4 } from "@utils/uuid";

// =============================================================================
// Types
// =============================================================================

/** Built-in column type names */
export type BuiltinColumnType =
  | "string"
  | "int"
  | "float"
  | "bool"
  | "date"
  | "datetime"
  | "uuid"
  | "email"
  | "name"
  | "firstName"
  | "lastName"
  | "word"
  | "sentence"
  | "paragraph"
  | "phone"
  | "url"
  | "ip"
  | "ipv6"
  | "hex"
  | "index"
  | "company"
  | "country"
  | "currency"
  | "percent"
  | "timestamp"
  | "city"
  | "zipCode"
  | "color"
  | "username"
  | "slug";

/** Every {@link BuiltinColumnType}, for checking a column definition. */
const BUILTIN_COLUMN_TYPES: ReadonlySet<string> = new Set<BuiltinColumnType>([
  "string",
  "int",
  "float",
  "bool",
  "date",
  "datetime",
  "uuid",
  "email",
  "name",
  "firstName",
  "lastName",
  "word",
  "sentence",
  "paragraph",
  "phone",
  "url",
  "ip",
  "ipv6",
  "hex",
  "index",
  "company",
  "country",
  "currency",
  "percent",
  "timestamp",
  "city",
  "zipCode",
  "color",
  "username",
  "slug"
]);

/** Custom generator function */
export type GeneratorFn = (context: GeneratorContext) => unknown;

/** Column definition */
export type ColumnDef = BuiltinColumnType | GeneratorFn | GeneratorColumnConfig;

/** Detailed column configuration */
export interface GeneratorColumnConfig {
  /** Column type or custom generator */
  type: BuiltinColumnType | GeneratorFn;
  /** Column name (for header) */
  name?: string;
  /** Minimum value (for int/float) */
  min?: number;
  /** Maximum value (for int/float) */
  max?: number;
  /** Length or max length (for string types) */
  length?: number;
  /** Possible values to pick from */
  values?: unknown[];
  /** Null probability (0-1) */
  nullable?: number;
  /** Date range start */
  dateFrom?: Date;
  /** Date range end */
  dateTo?: Date;
}

/** Context passed to generator functions */
export interface GeneratorContext {
  /** Current row index (0-based) */
  rowIndex: number;
  /** Current column index (0-based) */
  colIndex: number;
  /** Column name if defined */
  colName?: string;
  /** Random number generator (0-1) */
  random: () => number;
  /** Random integer in range */
  randomInt: (min: number, max: number) => number;
  /** Random float in range */
  randomFloat: (min: number, max: number, precision?: number) => number;
  /** Random item from array */
  randomPick: <T>(arr: readonly T[]) => T;
  /** Random string of given length */
  randomString: (length: number) => string;
  /** Random date in range */
  randomDate: (from?: Date, to?: Date) => Date;
  /** Random boolean with probability */
  randomBool: (probability?: number) => boolean;
}

/** Stop condition for generation */
export type StopCondition =
  | { rows: number }
  | { duration: number }
  | { until: Date | number }
  | { custom: (ctx: StopContext) => boolean };

/** Context for stop condition evaluation */
export interface StopContext {
  /** Number of rows generated so far */
  rowCount: number;
  /** Time elapsed since start (ms) */
  elapsed: number;
  /** Start time */
  startTime: number;
}

/** CSV generator options */
export interface CsvGenerateOptions {
  /** Column definitions (number for count, or array of definitions) */
  columns?: number | readonly ColumnDef[];
  /**
   * Number of rows to generate.
   * Use Infinity or -1 for unlimited (must use iterator API).
   */
  rows?: number;
  /** Seed for reproducible random data */
  seed?: number;
  /** Include header row */
  headers?: boolean | readonly string[];
  /** Field delimiter */
  delimiter?: string;
  /** Row delimiter (line ending) */
  lineEnding?: string;
  /**
   * Append string at end of output (e.g., "\n" for trailing newline).
   * Only applies to Csv.generate(), not iterators.
   */
  eof?: string;
  /**
   * Prepend UTF-8 BOM (Byte Order Mark) for Excel compatibility.
   * Only applies to Csv.generate(), not iterators.
   */
  bom?: boolean;
  /**
   * Output rows as objects with header keys instead of arrays.
   * Only applies to Csv.generateData().
   */
  objectMode?: boolean;
  /**
   * Transform function applied to each generated row.
   * Receives the row data and context, returns the transformed row.
   */
  transform?: (row: unknown[], context: { rowIndex: number; headers: string[] }) => unknown[];
  /**
   * Stop after duration milliseconds (for iterator APIs).
   * Takes precedence over rows if both specified.
   */
  duration?: number;
  /**
   * Stop at specific time (Date or timestamp).
   * Takes precedence over rows if specified.
   */
  until?: Date | number;
  /** Default string length */
  stringLength?: number;
  /** Default int range */
  intMin?: number;
  intMax?: number;
  /** Default float range */
  floatMin?: number;
  floatMax?: number;
  /** Default float precision */
  floatPrecision?: number;
  /**
   * Quoting strategy for CSV fields.
   * - 'auto': Quote only when necessary (default)
   * - 'always': Always quote all fields
   * - 'never': Never quote (may produce invalid CSV if data contains delimiters)
   */
  quote?: "auto" | "always" | "never";
  /**
   * Skip first N data rows (headers are not affected).
   * Useful for resuming generation or pagination.
   */
  skipRows?: number;
}

/** Result from generate function */
export interface CsvGenerateResult {
  /** Generated CSV string */
  csv: string;
  /** Header names */
  headers: string[];
  /** Data as 2D array */
  data: unknown[][];
}

// =============================================================================
// Random Generator (Mulberry32 PRNG)
// =============================================================================

/**
 * Create a seeded pseudo-random number generator using Mulberry32 algorithm.
 * Fast and produces good quality random numbers for non-cryptographic use.
 */
function createSeededRandom(seed: number): () => number {
  let state = seed >>> 0;
  return (): number => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// =============================================================================
// Built-in Data Generators
// =============================================================================

const ALPHA_LOWER = "abcdefghijklmnopqrstuvwxyz";
const ALPHA_NUM = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
const HEX_CHARS = "0123456789abcdef";

const FIRST_NAMES = [
  "James",
  "Mary",
  "John",
  "Patricia",
  "Robert",
  "Jennifer",
  "Michael",
  "Linda",
  "William",
  "Elizabeth",
  "David",
  "Barbara",
  "Richard",
  "Susan",
  "Joseph",
  "Jessica",
  "Thomas",
  "Sarah",
  "Charles",
  "Karen",
  "Emma",
  "Olivia",
  "Ava",
  "Sophia",
  "Isabella",
  "Mia",
  "Charlotte",
  "Amelia",
  "Harper",
  "Evelyn",
  "Liam",
  "Noah",
  "Oliver",
  "Elijah",
  "Lucas",
  "Mason",
  "Logan",
  "Alexander",
  "Ethan",
  "Jacob"
];

const LAST_NAMES = [
  "Smith",
  "Johnson",
  "Williams",
  "Brown",
  "Jones",
  "Garcia",
  "Miller",
  "Davis",
  "Rodriguez",
  "Martinez",
  "Hernandez",
  "Lopez",
  "Gonzalez",
  "Wilson",
  "Anderson",
  "Thomas",
  "Taylor",
  "Moore",
  "Jackson",
  "Martin",
  "Lee",
  "Perez",
  "Thompson",
  "White",
  "Harris",
  "Sanchez",
  "Clark",
  "Ramirez",
  "Lewis",
  "Robinson"
];

const WORDS = [
  "the",
  "be",
  "to",
  "of",
  "and",
  "a",
  "in",
  "that",
  "have",
  "it",
  "for",
  "not",
  "on",
  "with",
  "he",
  "as",
  "you",
  "do",
  "at",
  "this",
  "but",
  "his",
  "by",
  "from",
  "they",
  "we",
  "say",
  "her",
  "she",
  "or",
  "an",
  "will",
  "my",
  "one",
  "all",
  "would",
  "there",
  "their",
  "what",
  "so",
  "up",
  "out",
  "if",
  "about",
  "who",
  "get",
  "which",
  "go",
  "me",
  "when",
  "make",
  "can",
  "like",
  "time",
  "no",
  "just",
  "him",
  "know",
  "take",
  "people",
  "into",
  "year",
  "your",
  "good",
  "some",
  "could",
  "them",
  "see",
  "other",
  "than",
  "then",
  "now",
  "look",
  "only",
  "come",
  "its",
  "over",
  "think",
  "also",
  "back",
  "after",
  "use",
  "two",
  "how",
  "our",
  "work",
  "first",
  "well",
  "way",
  "even",
  "new",
  "want",
  "because",
  "any",
  "these",
  "give",
  "day",
  "most",
  "us"
];

const EMAIL_DOMAINS = [
  "gmail.com",
  "yahoo.com",
  "outlook.com",
  "hotmail.com",
  "example.com",
  "test.org",
  "company.io",
  "mail.net"
];

const URL_PREFIXES = ["https://www.", "https://", "http://www.", "http://"];
const URL_DOMAINS = ["example.com", "test.org", "demo.io", "sample.net", "mock.dev"];
const URL_PATHS = ["", "/about", "/products", "/services", "/contact", "/blog", "/api/v1"];

const COMPANIES = [
  "Acme Corp",
  "Globex Inc",
  "Initech",
  "Umbrella Corp",
  "Stark Industries",
  "Wayne Enterprises",
  "Cyberdyne Systems",
  "Aperture Science",
  "Massive Dynamic",
  "Tyrell Corp",
  "Weyland-Yutani",
  "Soylent Corp",
  "Oscorp",
  "LexCorp",
  "Wonka Industries"
];

const COUNTRIES = [
  "United States",
  "United Kingdom",
  "Canada",
  "Australia",
  "Germany",
  "France",
  "Japan",
  "China",
  "India",
  "Brazil",
  "Mexico",
  "Italy",
  "Spain",
  "South Korea",
  "Netherlands"
];

const CURRENCIES = ["USD", "EUR", "GBP", "JPY", "CNY", "CAD", "AUD", "CHF", "INR", "BRL"];

const CITIES = [
  "New York",
  "Los Angeles",
  "Chicago",
  "Houston",
  "Phoenix",
  "London",
  "Paris",
  "Tokyo",
  "Sydney",
  "Toronto",
  "Berlin",
  "Madrid",
  "Rome",
  "Seoul",
  "Mumbai"
];

const COLORS = [
  "red",
  "blue",
  "green",
  "yellow",
  "purple",
  "orange",
  "pink",
  "black",
  "white",
  "gray",
  "brown",
  "cyan",
  "magenta",
  "lime",
  "navy"
];

// =============================================================================
// Helper Functions
// =============================================================================

/**
 * Generate header names from column definitions
 */
function generateHeaderNames(
  colDefs: readonly ColumnDef[],
  providedHeaders?: boolean | readonly string[]
): string[] {
  const headerNames: string[] = [];

  if (providedHeaders === false) {
    return headerNames;
  }

  if (Array.isArray(providedHeaders)) {
    headerNames.push(...providedHeaders);
    // Pad if needed
    while (headerNames.length < colDefs.length) {
      headerNames.push(`col_${headerNames.length + 1}`);
    }
    return headerNames;
  }

  // Auto-generate headers
  for (let i = 0; i < colDefs.length; i++) {
    const col = colDefs[i];
    if (typeof col === "object" && "name" in col && col.name) {
      headerNames.push(col.name);
    } else if (typeof col === "string") {
      headerNames.push(`${col}_${i + 1}`);
    } else {
      headerNames.push(`col_${i + 1}`);
    }
  }

  return headerNames;
}

function checkGenerateOptions(options: CsvGenerateOptions, where: string): void {
  if (typeof options !== "object" || options === null || Array.isArray(options)) {
    throw new CsvArgumentError(
      "options",
      `${where}: options must be an options object, got ${options === null ? "null" : typeof options}`
    );
  }
  const fail = (option: string, expected: string, value: unknown): never => {
    throw new CsvArgumentError(
      `options.${option}`,
      `${where}: ${option} must be ${expected}, got ${typeof value === "string" ? JSON.stringify(value) : String(value)}`
    );
  };
  const integer = (option: keyof CsvGenerateOptions, min: number): void => {
    const value = options[option];
    if (
      value !== undefined &&
      (typeof value !== "number" || !Number.isInteger(value) || value < min)
    ) {
      fail(option, `an integer >= ${min}`, value);
    }
  };
  const finite = (option: keyof CsvGenerateOptions): void => {
    const value = options[option];
    if (value !== undefined && (typeof value !== "number" || !Number.isFinite(value))) {
      fail(option, "a finite number", value);
    }
  };
  if (typeof options.columns === "number") {
    integer("columns", 0);
  } else if (options.columns !== undefined && !Array.isArray(options.columns)) {
    fail("columns", "a column count or an array of column definitions", options.columns);
  } else if (options.columns !== undefined) {
    options.columns.forEach((def, i) => {
      const type =
        typeof def === "object" && def !== null ? (def as GeneratorColumnConfig).type : def;
      if (typeof type !== "function" && !BUILTIN_COLUMN_TYPES.has(type as string)) {
        fail(`columns[${i}]`, "a built-in column type, a function or { type }", type);
      }
    });
  }
  if (options.rows !== Infinity && options.rows !== -1) {
    integer("rows", 0);
  }
  for (const option of ["delimiter", "lineEnding"] as const) {
    const value = options[option];
    if (value !== undefined && (typeof value !== "string" || value === "")) {
      fail(option, "a non-empty string", value);
    }
  }
  if (options.eof !== undefined && typeof options.eof !== "string") {
    fail("eof", "a string", options.eof);
  }
  if (options.quote !== undefined && !["auto", "always", "never"].includes(options.quote)) {
    fail("quote", '"auto", "always" or "never"', options.quote);
  }
  finite("seed");
  integer("skipRows", 0);
  integer("stringLength", 0);
  // `toFixed` accepts 0 to 100 digits.
  integer("floatPrecision", 0);
  if ((options.floatPrecision ?? 0) > 100) {
    fail("floatPrecision", "an integer from 0 to 100", options.floatPrecision);
  }
  for (const option of ["intMin", "intMax", "floatMin", "floatMax"] as const) {
    finite(option);
  }
  finite("duration");
  if (
    options.until !== undefined &&
    !(options.until instanceof Date
      ? Number.isFinite(options.until.getTime())
      : Number.isFinite(options.until))
  ) {
    fail("until", "a valid Date or timestamp", options.until);
  }
  if (options.transform !== undefined && typeof options.transform !== "function") {
    fail("transform", "a function", options.transform);
  }
}

/**
 * Normalize column definitions
 */
function normalizeColumns(columns: number | readonly ColumnDef[]): readonly ColumnDef[] {
  return typeof columns === "number" ? Array(columns).fill("string") : columns;
}

// =============================================================================
// Generator Class
// =============================================================================

class CsvGenerator {
  private random: () => number;
  private seeded: boolean;
  private options: Required<
    Pick<
      CsvGenerateOptions,
      | "delimiter"
      | "lineEnding"
      | "stringLength"
      | "intMin"
      | "intMax"
      | "floatMin"
      | "floatMax"
      | "floatPrecision"
      | "quote"
    >
  >;

  constructor(seed?: number, options?: CsvGenerateOptions) {
    this.seeded = seed !== undefined;
    this.random = this.seeded ? createSeededRandom(seed!) : Math.random;
    this.options = {
      delimiter: options?.delimiter ?? ",",
      lineEnding: options?.lineEnding ?? "\n",
      stringLength: options?.stringLength ?? 10,
      intMin: options?.intMin ?? 0,
      intMax: options?.intMax ?? 10000,
      floatMin: options?.floatMin ?? 0,
      floatMax: options?.floatMax ?? 1000,
      floatPrecision: options?.floatPrecision ?? 2,
      quote: options?.quote ?? "auto"
    };
  }

  // ---------------------------------------------------------------------------
  // Utility Methods
  // ---------------------------------------------------------------------------

  randomInt(min: number, max: number): number {
    return Math.floor(this.random() * (max - min + 1)) + min;
  }

  randomFloat(min: number, max: number, precision: number): number {
    const value = this.random() * (max - min) + min;
    const factor = Math.pow(10, precision);
    return Math.round(value * factor) / factor;
  }

  randomPick<T>(arr: readonly T[]): T {
    return arr[Math.floor(this.random() * arr.length)];
  }

  randomString(length: number, chars: string = ALPHA_NUM): string {
    // Use array join for better performance with longer strings
    const charsLen = chars.length;
    const result = new Array(length);
    for (let i = 0; i < length; i++) {
      result[i] = chars[Math.floor(this.random() * charsLen)];
    }
    return result.join("");
  }

  randomBool(probability: number = 0.5): boolean {
    return this.random() < probability;
  }

  randomDate(from: Date, to: Date): Date {
    const fromTime = from.getTime();
    const toTime = to.getTime();
    return new Date(fromTime + this.random() * (toTime - fromTime));
  }

  // ---------------------------------------------------------------------------
  // Value Generators
  // ---------------------------------------------------------------------------

  generateValue(colDef: ColumnDef, ctx: GeneratorContext): unknown {
    // Handle function directly
    if (typeof colDef === "function") {
      return colDef(ctx);
    }

    // Handle string type name
    if (typeof colDef === "string") {
      return this.generateBuiltinType(colDef, {}, ctx);
    }

    // Handle config object
    const config = colDef as GeneratorColumnConfig;

    // Check nullable
    if (config.nullable !== undefined && this.random() < config.nullable) {
      return null;
    }

    // Check values array (enum-like)
    if (config.values && config.values.length > 0) {
      return this.randomPick(config.values);
    }

    // Handle custom function in config
    if (typeof config.type === "function") {
      return config.type(ctx);
    }

    return this.generateBuiltinType(config.type, config, ctx);
  }

  private generateBuiltinType(
    type: BuiltinColumnType,
    config: Partial<GeneratorColumnConfig>,
    ctx: GeneratorContext
  ): unknown {
    switch (type) {
      case "string":
        return this.randomString(config.length ?? this.options.stringLength);

      case "int":
        return this.randomInt(config.min ?? this.options.intMin, config.max ?? this.options.intMax);

      case "float":
        return this.randomFloat(
          config.min ?? this.options.floatMin,
          config.max ?? this.options.floatMax,
          this.options.floatPrecision
        );

      case "bool":
        return this.randomBool();

      case "date": {
        const from = config.dateFrom ?? new Date(2020, 0, 1);
        const to = config.dateTo ?? new Date(2025, 11, 31);
        const d = this.randomDate(from, to);
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
      }

      case "datetime": {
        const from = config.dateFrom ?? new Date(2020, 0, 1);
        const to = config.dateTo ?? new Date(2025, 11, 31);
        return this.randomDate(from, to).toISOString();
      }

      case "uuid":
        // Use seeded random for UUID when seeded
        if (this.seeded) {
          const bytes = new Uint8Array(16);
          for (let i = 0; i < 16; i++) {
            bytes[i] = Math.floor(this.random() * 256);
          }
          // Set version (4) and variant (10xx)
          bytes[6] = (bytes[6] & 0x0f) | 0x40;
          bytes[8] = (bytes[8] & 0x3f) | 0x80;
          const hex = (b: number) => b.toString(16).padStart(2, "0");
          return `${hex(bytes[0])}${hex(bytes[1])}${hex(bytes[2])}${hex(bytes[3])}-${hex(bytes[4])}${hex(bytes[5])}-${hex(bytes[6])}${hex(bytes[7])}-${hex(bytes[8])}${hex(bytes[9])}-${hex(bytes[10])}${hex(bytes[11])}${hex(bytes[12])}${hex(bytes[13])}${hex(bytes[14])}${hex(bytes[15])}`;
        }
        return uuidV4();

      case "email": {
        const name = this.randomString(config.length ?? 8, ALPHA_LOWER);
        const domain = this.randomPick(EMAIL_DOMAINS);
        return `${name}@${domain}`;
      }

      case "name": {
        const first = this.randomPick(FIRST_NAMES);
        const last = this.randomPick(LAST_NAMES);
        return `${first} ${last}`;
      }

      case "firstName":
        return this.randomPick(FIRST_NAMES);

      case "lastName":
        return this.randomPick(LAST_NAMES);

      case "word":
        return this.randomPick(WORDS);

      case "sentence": {
        const wordCount = config.length ?? this.randomInt(5, 12);
        const words: string[] = [];
        for (let i = 0; i < wordCount; i++) {
          words.push(this.randomPick(WORDS));
        }
        // Capitalize first letter
        words[0] = words[0].charAt(0).toUpperCase() + words[0].slice(1);
        return words.join(" ") + ".";
      }

      case "paragraph": {
        const sentenceCount = config.length ?? this.randomInt(3, 6);
        const sentences: string[] = [];
        for (let i = 0; i < sentenceCount; i++) {
          const wordCount = this.randomInt(5, 12);
          const words: string[] = [];
          for (let j = 0; j < wordCount; j++) {
            words.push(this.randomPick(WORDS));
          }
          words[0] = words[0].charAt(0).toUpperCase() + words[0].slice(1);
          sentences.push(words.join(" ") + ".");
        }
        return sentences.join(" ");
      }

      case "phone": {
        // Generate phone in format: +1-XXX-XXX-XXXX
        const area = this.randomInt(200, 999);
        const prefix = this.randomInt(200, 999);
        const line = this.randomInt(1000, 9999);
        return `+1-${area}-${prefix}-${line}`;
      }

      case "url": {
        const prefix = this.randomPick(URL_PREFIXES);
        const domain = this.randomPick(URL_DOMAINS);
        const path = this.randomPick(URL_PATHS);
        return `${prefix}${domain}${path}`;
      }

      case "ip": {
        // IPv4 address
        const octets = [
          this.randomInt(1, 255),
          this.randomInt(0, 255),
          this.randomInt(0, 255),
          this.randomInt(1, 254)
        ];
        return octets.join(".");
      }

      case "ipv6": {
        // IPv6 address (simplified)
        const groups: string[] = [];
        for (let i = 0; i < 8; i++) {
          groups.push(this.randomString(4, HEX_CHARS));
        }
        return groups.join(":");
      }

      case "hex":
        return this.randomString(config.length ?? 8, HEX_CHARS);

      case "index":
        return ctx.rowIndex;

      case "company":
        return this.randomPick(COMPANIES);

      case "country":
        return this.randomPick(COUNTRIES);

      case "currency": {
        const amount = this.randomFloat(
          config.min ?? 0,
          config.max ?? 10000,
          this.options.floatPrecision
        );
        const currency = this.randomPick(CURRENCIES);
        return `${currency} ${amount.toFixed(2)}`;
      }

      case "percent":
        return this.randomFloat(config.min ?? 0, config.max ?? 100, 1) + "%";

      case "timestamp": {
        // Use a fixed base when seeded for reproducible output
        const base = this.seeded ? 1704067200000 : Date.now(); // 2024-01-01T00:00:00Z when seeded
        return base + this.randomInt(0, config.max ?? 86400000);
      }

      case "city":
        return this.randomPick(CITIES);

      case "zipCode": {
        // US-style ZIP code
        const zip = this.randomInt(10000, 99999);
        return String(zip);
      }

      case "color":
        return this.randomPick(COLORS);

      case "username": {
        const adjectives = [
          "cool",
          "super",
          "mega",
          "ultra",
          "pro",
          "epic",
          "fast",
          "dark",
          "light",
          "wild"
        ];
        const adj = this.randomPick(adjectives);
        const name = this.randomPick(FIRST_NAMES).toLowerCase();
        const num = this.randomInt(1, 999);
        return `${adj}_${name}${num}`;
      }

      case "slug": {
        const words: string[] = [];
        const count = config.length ?? this.randomInt(2, 4);
        for (let i = 0; i < count; i++) {
          words.push(this.randomPick(WORDS));
        }
        return words.join("-");
      }

      default:
        return this.randomString(this.options.stringLength);
    }
  }

  // ---------------------------------------------------------------------------
  // Context Factory
  // ---------------------------------------------------------------------------

  createContext(rowIndex: number, colIndex: number, colName?: string): GeneratorContext {
    return {
      rowIndex,
      colIndex,
      colName,
      random: () => this.random(),
      randomInt: (min, max) => this.randomInt(min, max),
      randomFloat: (min, max, precision = 2) => this.randomFloat(min, max, precision),
      randomPick: <T>(arr: readonly T[]) => this.randomPick(arr),
      randomString: length => this.randomString(length),
      randomDate: (from, to) =>
        this.randomDate(from ?? new Date(2020, 0, 1), to ?? new Date(2025, 11, 31)),
      randomBool: (probability = 0.5) => this.randomBool(probability)
    };
  }

  // ---------------------------------------------------------------------------
  // Row Generation
  // ---------------------------------------------------------------------------

  /**
   * Generate a single row of data
   */
  generateRow(
    colDefs: readonly ColumnDef[],
    rowIndex: number,
    headerNames: string[],
    transform?: CsvGenerateOptions["transform"]
  ): unknown[] {
    const row: unknown[] = [];
    for (let colIdx = 0; colIdx < colDefs.length; colIdx++) {
      const ctx = this.createContext(rowIndex, colIdx, headerNames[colIdx]);
      row.push(this.generateValue(colDefs[colIdx], ctx));
    }
    if (transform) {
      return transform(row, { rowIndex, headers: headerNames });
    }
    return row;
  }

  // ---------------------------------------------------------------------------
  // CSV Formatting
  // ---------------------------------------------------------------------------

  formatField(value: unknown): string {
    if (value === null || value === undefined) {
      return this.options.quote === "always" ? '""' : "";
    }

    const str = String(value);

    // Never quote mode - fastest but may produce invalid CSV
    if (this.options.quote === "never") {
      return str;
    }

    // Always quote mode
    if (this.options.quote === "always") {
      return '"' + str.replace(/"/g, '""') + '"';
    }

    // Auto mode - quote only when necessary
    if (
      !str.includes(this.options.delimiter) &&
      !str.includes('"') &&
      !str.includes("\n") &&
      !str.includes("\r")
    ) {
      return str;
    }

    return '"' + str.replace(/"/g, '""') + '"';
  }

  formatRow(values: unknown[]): string {
    return values.map(v => this.formatField(v)).join(this.options.delimiter);
  }
}

// =============================================================================
// Stop Condition Helper
// =============================================================================

/**
 * Create a stop condition checker from options.
 * Returns a function that returns true when generation should stop.
 */
function createStopChecker(options: CsvGenerateOptions): (ctx: StopContext) => boolean {
  const { rows = 10, duration, until } = options;

  // Duration-based stop
  if (duration !== undefined && duration > 0) {
    return (ctx: StopContext) => ctx.elapsed >= duration;
  }

  // Until-based stop (specific time)
  if (until !== undefined) {
    const endTime = until instanceof Date ? until.getTime() : until;
    return () => Date.now() >= endTime;
  }

  // Row-based stop (default)
  // -1 or Infinity means unlimited
  if (rows === -1 || rows === Infinity) {
    return () => false; // Never stop
  }

  return (ctx: StopContext) => ctx.rowCount >= rows;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Generate CSV data synchronously.
 *
 * @example
 * ```ts
 * // Simple: 5 columns, 100 rows
 * const { csv } = csvGenerate({ columns: 5, rows: 100 });
 *
 * // With column types
 * const { csv, data } = csvGenerate({
 *   columns: ['name', 'email', 'int', 'bool', 'date'],
 *   rows: 50,
 *   headers: ['Name', 'Email', 'Age', 'Active', 'JoinDate']
 * });
 *
 * // With custom generators
 * const { csv } = csvGenerate({
 *   columns: [
 *     { type: 'int', min: 18, max: 65, name: 'age' },
 *     { type: 'float', min: 0, max: 100, name: 'score' },
 *     (ctx) => `row-${ctx.rowIndex}`,
 *   ],
 *   rows: 100,
 *   seed: 12345  // Reproducible
 * });
 * ```
 */
export function csvGenerate(options: CsvGenerateOptions = {}): CsvGenerateResult {
  checkGenerateOptions(options, "Csv.generate");
  const { columns = 5, rows = 10, seed, headers, eof, bom, transform, skipRows = 0 } = options;

  // Prevent infinite generation in sync API
  if (rows === -1 || rows === Infinity) {
    throw new CsvArgumentError(
      "options.rows",
      "Unlimited generation (rows: -1 or Infinity) is not supported in Csv.generate(). " +
        "Use the Csv.generateRows() iterator instead."
    );
  }

  const generator = new CsvGenerator(seed, options);
  const colDefs = normalizeColumns(columns);
  const headerNames = generateHeaderNames(colDefs, headers);

  // Generate data using helper (skip first N rows if specified)
  const data: unknown[][] = [];
  const totalRows = rows + skipRows;
  for (let rowIdx = 0; rowIdx < totalRows; rowIdx++) {
    const row = generator.generateRow(colDefs, rowIdx, headerNames, transform);
    if (rowIdx >= skipRows) {
      data.push(row);
    }
  }

  // Build CSV string
  const lines: string[] = [];

  // Add header row
  if (headers !== false && headerNames.length > 0) {
    lines.push(generator.formatRow(headerNames));
  }

  // Add data rows
  for (const row of data) {
    lines.push(generator.formatRow(row));
  }

  const lineEnding = options.lineEnding ?? "\n";
  let csv = lines.join(lineEnding);

  // Prepend BOM if specified
  if (bom) {
    csv = "\uFEFF" + csv;
  }

  // Append eof if specified
  if (eof !== undefined) {
    csv += eof;
  }

  return { csv, headers: headerNames, data };
}

/**
 * Generate CSV rows as an iterator (memory efficient for large datasets).
 *
 * Supports unlimited generation with stop conditions:
 * - `rows: -1` or `rows: Infinity` for unlimited
 * - `duration: 5000` to generate for 5 seconds
 * - `until: new Date('2024-12-31')` to generate until specific time
 *
 * @example
 * ```ts
 * // Fixed number of rows
 * for (const row of csvGenerateRows({ columns: 5, rows: 1000000 })) {
 *   process.stdout.write(row + '\n');
 * }
 *
 * // Generate for 5 seconds
 * for (const row of csvGenerateRows({ columns: 3, duration: 5000 })) {
 *   console.log(row);
 * }
 *
 * // Unlimited with manual break
 * for (const row of csvGenerateRows({ columns: 3, rows: -1 })) {
 *   if (someCondition) break;
 *   console.log(row);
 * }
 * ```
 */
export function* csvGenerateRows(
  options: CsvGenerateOptions = {}
): Generator<string, void, undefined> {
  checkGenerateOptions(options, "Csv.generateRows");
  const { columns = 5, seed, headers, transform, skipRows = 0 } = options;

  const generator = new CsvGenerator(seed, options);
  const shouldStop = createStopChecker(options);
  const startTime = Date.now();
  const colDefs = normalizeColumns(columns);
  const headerNames = generateHeaderNames(colDefs, headers);

  // Yield header row
  if (headers !== false && headerNames.length > 0) {
    yield generator.formatRow(headerNames);
  }

  // Yield data rows using helper
  let rowIdx = 0;
  let yieldedCount = 0;
  while (true) {
    const stopCtx: StopContext = {
      rowCount: yieldedCount,
      elapsed: Date.now() - startTime,
      startTime
    };

    if (shouldStop(stopCtx)) {
      break;
    }

    const row = generator.generateRow(colDefs, rowIdx, headerNames, transform);
    rowIdx++;

    // Skip first N rows if specified
    if (rowIdx <= skipRows) {
      continue;
    }

    yield generator.formatRow(row);
    yieldedCount++;
  }
}

/**
 * Generate CSV data as an async generator (useful for streaming with delays).
 *
 * Supports all stop conditions from csvGenerateRows plus delay between rows.
 *
 * @example
 * ```ts
 * // Generate with delay between rows
 * for await (const row of csvGenerateAsync({ columns: 5, rows: 100, delay: 10 })) {
 *   console.log(row);
 * }
 *
 * // Generate for 5 seconds with 100ms delay
 * for await (const row of csvGenerateAsync({ columns: 3, duration: 5000, delay: 100 })) {
 *   console.log(row);
 * }
 * ```
 */
export async function* csvGenerateAsync(
  options: CsvGenerateOptions & { delay?: number } = {}
): AsyncGenerator<string, void, undefined> {
  checkGenerateOptions(options, "Csv.generateRowsAsync");
  const { columns = 5, seed, headers, delay = 0, transform, skipRows = 0 } = options;

  const generator = new CsvGenerator(seed, options);
  const shouldStop = createStopChecker(options);
  const startTime = Date.now();
  const colDefs = normalizeColumns(columns);
  const headerNames = generateHeaderNames(colDefs, headers);

  // Yield header row
  if (headers !== false && headerNames.length > 0) {
    yield generator.formatRow(headerNames);
    if (delay > 0) {
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }

  // Yield data rows using helper
  let rowIdx = 0;
  let yieldedCount = 0;
  while (true) {
    const stopCtx: StopContext = {
      rowCount: yieldedCount,
      elapsed: Date.now() - startTime,
      startTime
    };

    if (shouldStop(stopCtx)) {
      break;
    }

    const row = generator.generateRow(colDefs, rowIdx, headerNames, transform);
    rowIdx++;

    // Skip first N rows if specified
    if (rowIdx <= skipRows) {
      continue;
    }

    yield generator.formatRow(row);
    yieldedCount++;

    if (delay > 0) {
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
}

/**
 * Generate raw data rows (without CSV formatting).
 *
 * @example
 * ```ts
 * const rows = csvGenerateData({
 *   columns: ['name', 'int', 'email'],
 *   rows: 10
 * });
 * // Returns: [['John Smith', 42, 'abc@gmail.com'], ...]
 *
 * // With objectMode
 * const objects = csvGenerateData({
 *   columns: [{ type: 'name', name: 'fullName' }, { type: 'int', name: 'age' }],
 *   rows: 10,
 *   objectMode: true
 * });
 * // Returns: [{ fullName: 'John Smith', age: 42 }, ...]
 * ```
 */
export function csvGenerateData<T extends CsvGenerateOptions>(
  options?: T
): T extends { objectMode: true } ? Record<string, unknown>[] : unknown[][] {
  const opts = options ?? ({} as T);
  checkGenerateOptions(opts, "Csv.generateData");
  const { columns = 5, rows = 10, seed, objectMode, transform, skipRows = 0 } = opts;

  // Prevent infinite generation in sync API
  if (rows === -1 || rows === Infinity) {
    throw new CsvArgumentError(
      "options.rows",
      "Unlimited generation (rows: -1 or Infinity) is not supported in Csv.generateData(). " +
        "Use the Csv.generateRows() iterator instead."
    );
  }

  const generator = new CsvGenerator(seed, opts);
  const colDefs = normalizeColumns(columns);
  const headerNames = generateHeaderNames(colDefs, opts.headers);
  const totalRows = rows + skipRows;

  // Generate data using helper
  if (objectMode) {
    const data: Record<string, unknown>[] = [];
    for (let rowIdx = 0; rowIdx < totalRows; rowIdx++) {
      const row = generator.generateRow(colDefs, rowIdx, headerNames, transform);
      if (rowIdx >= skipRows) {
        const obj: Record<string, unknown> = {};
        for (let i = 0; i < headerNames.length; i++) {
          obj[headerNames[i]] = row[i];
        }
        data.push(obj);
      }
    }
    return data as T extends { objectMode: true } ? Record<string, unknown>[] : unknown[][];
  }

  const data: unknown[][] = [];
  for (let rowIdx = 0; rowIdx < totalRows; rowIdx++) {
    const row = generator.generateRow(colDefs, rowIdx, headerNames, transform);
    if (rowIdx >= skipRows) {
      data.push(row);
    }
  }
  return data as T extends { objectMode: true } ? Record<string, unknown>[] : unknown[][];
}

/**
 * Create a reusable generator instance with preset configuration.
 *
 * @example
 * ```ts
 * const gen = createCsvGenerator({
 *   columns: ['name', 'email', { type: 'int', min: 18, max: 99 }],
 *   seed: 42
 * });
 *
 * // Generate multiple batches with same config
 * const batch1 = gen.generate(100);
 * const batch2 = gen.generate(100);
 *
 * // Generate for 5 seconds
 * for (const row of gen.rows({ duration: 5000 })) {
 *   console.log(row);
 * }
 * ```
 */
export function createCsvGenerator(baseOptions: CsvGenerateOptions = {}) {
  checkGenerateOptions(baseOptions, "Csv.createGenerator");
  return {
    /**
     * Generate CSV synchronously
     */
    generate(rowsOrOptions?: number | Partial<CsvGenerateOptions>): CsvGenerateResult {
      const overrides =
        typeof rowsOrOptions === "number" ? { rows: rowsOrOptions } : (rowsOrOptions ?? {});
      return csvGenerate({ ...baseOptions, ...overrides });
    },

    /**
     * Generate CSV rows as iterator
     */
    *rows(
      rowsOrOptions?: number | Partial<CsvGenerateOptions>
    ): Generator<string, void, undefined> {
      const overrides =
        typeof rowsOrOptions === "number" ? { rows: rowsOrOptions } : (rowsOrOptions ?? {});
      yield* csvGenerateRows({ ...baseOptions, ...overrides });
    },

    /**
     * Generate raw data array
     */
    data<O extends Partial<CsvGenerateOptions>>(
      rowsOrOptions?: number | O
    ): O extends { objectMode: true } ? Record<string, unknown>[] : unknown[][] {
      const overrides =
        typeof rowsOrOptions === "number" ? { rows: rowsOrOptions } : (rowsOrOptions ?? {});
      return csvGenerateData({ ...baseOptions, ...overrides }) as O extends {
        objectMode: true;
      }
        ? Record<string, unknown>[]
        : unknown[][];
    },

    /**
     * Generate CSV rows as async iterator
     */
    async *asyncRows(
      options?: Partial<CsvGenerateOptions> & { delay?: number }
    ): AsyncGenerator<string, void, undefined> {
      yield* csvGenerateAsync({ ...baseOptions, ...options });
    }
  };
}

# 🚀 From Frustration to Innovation: My Development Track

> *"Each barrier is no more than a target to tear down."*

![](https://github-readme-stats.vercel.app/api?username=ws-garcia&theme=blue_navy&hide_border=false&include_all_commits=true&count_private=false)<br/>
![](https://nirzak-streak-stats.vercel.app/?user=ws-garcia&theme=blue_navy&hide_border=false)<br/>
![](https://github-readme-stats.vercel.app/api/top-langs/?username=ws-garcia&theme=blue_navy&hide_border=false&include_all_commits=true&count_private=true&layout=compact)


This document traces how a single practical need — parsing CSV files correctly in VBA — grew into a body of work spanning a complete scripting language, a peer-reviewed research paper, and tools being adopted by developers worldwide. Every project here started with a problem someone said was already solved or not worth solving.

---

## 📖 Table of Contents

- [Case 1: VBA CSV Interface — The Origin](#-case-1-vba-csv-interface--the-origin)
- [Case 2: CSV Dialect Detection](#-case-2-csv-dialect-detection)
- [Case 3: Expression Evaluation in VBA](#-case-3-expression-evaluation-in-vba)
- [Case 4: Modernizing VBA Development](#-case-4-modernizing-vba-development)
- [The Pattern](#-the-pattern-a-unified-development-philosophy)
- [Impact Summary](#-impact-summary)

---

## 📦 Case 1: VBA CSV Interface — The Origin

### 🔍 The Root Cause

CSV files are the universal lingua franca of data exchange — deceptively simple in concept, yet governed by no single globally accepted standard. [RFC-4180](https://www.rfc-editor.org/rfc/rfc4180), proposed by The Internet Society, provides a well-formed specification, but real-world CSV files deviate from it constantly: non-standard text qualifiers, embedded line breaks, locale-specific decimal separators, and escaped delimiters within fields.

I discovered that **Microsoft Excel is not fully compliant with RFC-4180**, and it provides no built-in mechanism to dump CSV content directly into a VBA array. Excel's legacy import tools, Power Query, and even Power BI each implement their own CSV parsing logic, and all of them fail on files that deviate from their narrow assumptions.

Consider a CSV file that uses the apostrophe (`'`) as a text qualifier instead of double quotes:

```
'This CSV file uses [CrLf]the Apostrophe char'	as fields delimiter
Power Query is not able to handle this kind	'of text qualifier, neither is Power BI.'
In the line bellow	we are using escapable fields
'My name is Jhon, the unique, Doe'	'my wife's name is 'Jane, wonderful, Doe'.'
```

This file has **4 records** and **2 fields** per record. The first field of the first record contains an embedded line break. The fourth record contains embedded commas and escaped apostrophes within qualified fields. Any compliant parser should handle this without difficulty.

### 🌐 Community Available Solutions

- **Excel's From Text (Legacy)**: Produced 5 records and 4 fields — ❌ incorrect on both counts. The embedded line break split a qualified field into a new row, and internal commas created phantom columns.

- **Power Query**: Produced 5 records and 6 fields — ❌ even worse. It could not handle the apostrophe text qualifier at all, fragmenting qualified fields across multiple columns and rows.

- **Power BI**: Suffered the same limitations. As one Power BI Community member documented, users were forced to *erase all quotation marks from the source data* or replace them with another character before import — a destructive workaround that corrupts the original data.

- **Community sentiment** 😒: When I first presented this problem and my proposed solution on [r/vba](https://www.reddit.com/r/vba/) and [r/excel](https://www.reddit.com/r/excel/), the reception was skeptical. One commenter dismissed the effort: *"I'm sure your tool fills some specific use case but I don't think you've done a very good job of explaining/selling it."*

| API | Output Records | Output Fields | Correct? |
|-----|:---:|:---:|:---:|
| From Text (Legacy) | 5 | 4 | ❌ |
| Power Query | 5 | 6 | ❌ |
| **CSV Interface** | **4** | **2** | ✅ |

### ✅ My Solution: VBA CSV Interface

I built a **complete, RFC-4180-compliant CSV parser implemented entirely in VBA** — with the flexibility to handle the non-standard dialects that real-world files actually use.

[**🔗 VBA CSV Interface on GitHub**](https://github.com/ws-garcia/VBA-CSV-interface)

What it provides:

- Full RFC-4180 compliance with configurable text qualifiers, delimiters, and escape characters
- Direct import of CSV data into VBA arrays — the capability Excel itself lacks
- Correct handling of embedded line breaks, escaped qualifiers, and multi-character delimiters
- A clean, minimal API:

```vb
Sub JoinTwoCSV()
    Dim csv As CSVinterface
    
    '@--------------------------------------------------
    ' Import data from CSV files
    Set csv = New CSVinterface
    With csv
        .parseConfig.delimitersGuessing = True
        .parseConfig.path = Environ("USERPROFILE") & "\Desktop\Sales details.csv"
        .ImportFromCSV .parseConfig
    End With
    
    '@--------------------------------------------------
    ' Write the result in a spreadsheet.
    csv.DumpToSheet DataSource:=csv.Items
End Sub
```

CSV Interface became the foundation from which every subsequent project emerged. The need for dialect detection led to CSVsniffer. The need for user-defined expressions led to VBA Expressions. And the architectural demands of those tools eventually led to ASF itself.

### 🧠 Mindset: Challenge the Platform, Not the User

Everyone assumed Excel's CSV handling was adequate because Excel is a Microsoft product and CSV is a "simple" format. I tested that assumption empirically and found it false — not in obscure edge cases, but in practical scenarios that real users encounter every day.

When told my tool was poorly explained or unnecessarily narrow, I didn't argue — I showed a side-by-side comparison where CSV Interface produced the correct output and every Microsoft tool failed. **Let the results do the talking.** That became my pattern.

The deeper insight: a CSV parser is not a trivial utility. It's the gateway between raw data and structured computation. Getting it right is not a "specific use case" — it's the foundation for everything that follows.

---

## 🔬 Case 2: CSV Dialect Detection

### 🔍 The Root Cause

As CSV Interface grew, I needed it to automatically detect the dialect of unknown CSV files — what delimiter is used? What's the text qualifier? What's the line terminator? Real-world files create genuine ambiguity. A file like:

```
Prüfung1;Prüfung2;Prüfung
1,5;33,33;15,55
```

Is the comma a delimiter or a decimal separator? Is the semicolon structural or part of a value? The existing tools treated dialect detection as a **character-level pattern matching** problem, when it is fundamentally a **structural inference** problem.

### 🌐 Community Available Solutions

- **Python's `csv.Sniffer`**: The standard library solution. Uses frequency analysis with rudimentary heuristics. Weighted F1 score ≈ 0.80, with a catastrophic **19.83% failure ratio** — it crashes or refuses to answer for nearly one in five files. Even a CPython core developer, Serhiy Storchaka, [acknowledged its fundamental flaws](https://discuss.python.org/t/rewrite-csv-sniffer/92652/18): *"There are many many issues with the current sniffer, it needs a complete rewrite."*

- **[CleverCSV](https://github.com/alan-turing-institute/CleverCSV)** (Alan Turing Institute): More robust at F1 ≈ 0.84, but still limited by its dialect evaluation approach and carrying a 7.99% failure ratio.

- **[DuckDB](https://duckdb.org/) `sniff_csv`**: Surprisingly strong at F1 ≈ 0.88, but requires reading 100 records and still falls short on messy datasets.

- **Community sentiment** 😒: When I asked the [r/vba](https://www.reddit.com/r/vba/) community for complex CSV files to test my delimiter guesser, the [stdLambda](https://github.com/sancarn/stdVBA/tree/master/src/stdLambda) author (sancarn) responded: *"Interesting but then becomes somewhat pointless."*

### ✅ My Solution: The Table Uniformity Method

Instead of asking *"which character appears most consistently?"*, I asked: **"which candidate delimiter produces the most uniform table?"**

I reframed the entire problem from pattern matching to structural quality measurement. The method evaluates each candidate delimiter by parsing the file under that hypothesis and measuring the resulting table's uniformity — using metrics including Mean Absolute Deviation and Shannon Entropy (MADSE variant). A correct delimiter produces consistent column counts and coherent data types; a wrong one produces structural chaos.

[**🔗 CSVsniffer on GitHub**](https://github.com/ws-garcia/CSVsniffer)

The methodology was published as a peer-reviewed paper in SAGE's *Data Science* journal:

> 📄 [**"Detecting CSV File Dialects by Table Uniformity Measurement and Data Type Inference"**](https://journals.sagepub.com/doi/10.3233/DS-240062)
> [![DOI](https://zenodo.org/badge/DOI/10.3233/DS-240062.svg)](https://doi.org/10.3233/DS-240062)

The results speak for themselves:

| Metric | CSVsniffer MADSE | CleverCSV | csv.Sniffer |
|--------|:---:|:---:|:---:|
| Weighted F1 Score | **0.9378** | 0.8425 | 0.8049 |
| Average Failure Ratio | **2.86%** | 7.99% | 19.83% |
| Reliability Factor | **87.80%** | 72.96% | 67.54% |

### 🌍 Global Impact

The Table Uniformity method is now inspiring developers and communities worldwide:

🦀 **Rust — [CSV Nose](https://github.com/jqnatividad/csv-nose)**: Joel Natividad (maintainer of [qsv](https://github.com/dathere/qsv), a high-performance CSV toolkit) ported the Table Uniformity method to Rust as a drop-in replacement for qsv's sniffer. The implementation directly cites my paper, achieves **99.55% accuracy** on the W3C-CSVW test suite, and is published on [crates.io](https://crates.io/crates/csv-nose). It uses [CSVsniffer's test datasets](https://github.com/ws-garcia/CSVsniffer) as its benchmark suite — the same datasets I curated from sources like the [Pollock framework](https://github.com/HPI-Information-Systems/Pollock), [W3C CSVW](https://github.com/w3c/csvw/tree/gh-pages), and the [CSV Wrangling research](https://github.com/alan-turing-institute/CSV_Wrangling).

🐍 **Python stdlib discussion — [Rewrite CSV Sniffer](https://discuss.python.org/t/rewrite-csv-sniffer/92652)**: In May 2025, Andy Terrel (CPython contributor) opened a discussion on Python's official forum proposing to rewrite `csv.Sniffer` using the Table Uniformity algorithm. He stated: *"my research keeps leading to cleverCSV and it's table uniformity algorithms for sniffing CSV formats"*, directly linking to [CSVsniffer's codebase](https://github.com/ws-garcia/CSVsniffer). The thread gained traction with CPython core developers:

- **Serhiy Storchaka** (core dev) independently described the same approach: *"we should just try to parse the file with many different parameters simultaneously, and choose the one that does not fail and produces the most credible looking data."*
- **Andy Terrel** confirmed: *"That's basically what the [table uniformity method](https://journals.sagepub.com/doi/10.3233/DS-240062) does"* — and committed to preparing an implementation.
- **Stephen Rosen** (core dev): *"I am in favor of you doing this work."*
- Andy Terrel also noted: *"While CleverCSV seems to be a winner in many ways, it's sniffing ability is still behind the state of the art that the Table Uniformity Algorithm is showing."*

The discussion revealed a telling detail: Python's `csv` module expert is listed as **inactive**, the existing sniffer has only **15 test samples**, and even Pandas internally inherits `csv.Sniffer`'s limitations. The community acknowledges this is a dead battery that needs replacing — and my research is leading the way.

### 🧠 Mindset: The Empiricist Who Reframes the Question

The community dismissed my problem because they were thinking about it wrong. Character-frequency heuristics *are* somewhat pointless — on that, sancarn was correct. But I didn't accept the community's framing. Instead of trying to build a better character counter, I asked a deeper question: **what defines a correct parse?** The answer — table uniformity — shifted the entire domain from signal processing to structural inference.

The dismissive response didn't discourage me; it confirmed that the existing paradigm was a dead end. The willingness to pursue a "pointless" idea until it yielded a peer-reviewed publication and global adoption is something I'm proud of. I trust empirical results over community consensus.

---

## ⚡ Case 3: Expression Evaluation in VBA

### 🔍 The Root Cause

VBA lacks native support for evaluating mathematical expressions from strings at runtime. This was a critical limitation for CSV Interface — I needed users to define computed fields, filters, and transformations dynamically. Without string-based expression evaluation, every computation has to be hard-coded at design time, killing the flexibility that makes data tools genuinely useful.

### 🌐 Community Available Solutions

- **Volpi MathParser**: A legacy expression evaluator for VBA. I contacted the original developer, Michael Ruder, and his response was not encouraging: *"I did not use the MathParser for many years and I think I never used the last version 4.x."* A dead-end dependency with known corner-case failures.

- **[stdLambda](https://github.com/sancarn/stdVBA/tree/master/src/stdLambda)** (by sancarn): A well-crafted functional programming library that correctly handled the corner cases where Volpi's parser failed. I called it "fabulous" — and I meant it. But I found two fundamental limitations:

  1. **⏱️ The ~50x performance cliff**: stdLambda was efficient for raw arithmetic but showed approximately **50 times lower performance** when evaluating built-in math functions (`sin`, `cos`, `tan`, `abs`). This wasn't a tuning issue — it was an architectural limitation in how function calls were dispatched.

  2. **🔢 Positional variable binding** (`$1`, `$2`, `$3`...): Variables had to be passed in a pre-established order. This works for simple one-off lambdas, but it's fundamentally incompatible with **programmatic variable assignment** — where another library (like CSV Interface) needs to inject variable values based on user-provided expressions with arbitrary variable names.

  Even sancarn [acknowledged the limitation](https://www.reddit.com/r/vba/comments/sxf0g2/evaluating_expressions_from_strings_in_vba/), offering three separate workarounds: dictionary-based row access (`$1.Age`), proposed syntax sugar (`#'Age'`), and global variable binding (`BindGlobal`). Three workarounds for one use case confirmed the architectural gap. As he noted: *"It definitely is unfortunate that accessing and editing the locals table of stdLambda is non-trivial. Definitely something for the library to improve upon."*

### ✅ My Solution: VBA Expressions → ASF

What began as a search for an expression evaluator for CSV Interface became a complete scripting language.

**Phase 1 — VBA Expressions**: I built a mathematical expression evaluator with its own grammar and a non-traditional approach to parsing infix expressions. Key innovations:

- Named variables with order-independent assignment: `Run("x = 1; y = 2; z = 3")`
- Arrays defined by text strings, constants, floating-point notation, and Boolean logic
- "Analyze once, evaluate many" architecture — parse once, evaluate repeatedly with different values

**Phase 2 — [ASF (Advanced Scripting Framework)](https://github.com/ws-garcia/ASF)**: The evaluator's architecture naturally extended into a full JavaScript-like scripting language implemented entirely in VBA. ASF v3.0 features:

- A complete **parser, compiler, and virtual machine**
- First-class functions with closures and lexical scoping
- Class-based OOP with inheritance and `super()` constructor calls
- A module system with import/export and `.vas` file support
- Spread/rest operators, object member methods (`keys`, `values`, `forEach`, `map`, `filter`)
- Performance-optimized architecture using reusable `ASF_Map` instances (5-10x gains)
- Full VS Code tooling: syntax highlighting, diagnostics, hover documentation, code completion, 181 snippets
- Runtime placeholders (`$1`, `$2`, `$3`) for API parity with stdLambda — while maintaining the superior named-variable approach

ASF benchmarks confirm it significantly outperforms stdLambda on mathematical functions and functional transformations — the exact use cases where stdLambda exhibited its 50x performance cliff.

### 🧠 Mindset: Follow the Architecture, Not the Scope

My initial goal was modest: integrate an expression evaluator into CSV Interface. But each limitation pointed to a deeper architectural need, and I followed that thread to its logical conclusion.

stdLambda's positional variables weren't a bug to patch — they were a symptom of a design that treated lambdas as isolated computations rather than components in a larger system. Fixing the variable model required rethinking expression evaluation from the ground up. And once you have a proper parser, variable scope, and evaluation engine, you're one step from closures; one step from closures to first-class functions; one step from there to a complete language.

This isn't scope creep — it's **scope discovery**. As I wrote at the time: *"although the initial scope of the project was far exceeded, the level of difficulty in evaluating mathematical expressions without using the traditional method is overwhelming."* I chose the overwhelming path because it was the correct one.

---

## 🏗️ Case 4: Modernizing VBA Development

### 🔍 The Root Cause

VBA is a language frozen in time. No closures, no first-class functions, no modern iteration patterns, no module systems, no real OOP beyond basic class modules. Millions of enterprise Excel workbooks depend on VBA, but the language hasn't evolved in decades. Developers in this ecosystem face a choice: accept VBA's limitations, migrate to a different platform (often impractical), or find a way to bring modern capabilities into the existing environment.

The fundamental tension: **language capability** vs. **ecosystem lock-in**. VBA isn't going away, but it's inadequate for the complexity of modern business logic.

### 🌐 Community Available Solutions

- **[twinBASIC](https://twinbasic.com/)**: A modernization effort replacing VBA with a compatible but updated BASIC dialect. Requires adopting a new development environment and compiler toolchain — not always feasible in locked-down enterprise environments.

- **stdLambda and the [std library family](https://github.com/sancarn/stdVBA)**: Functional programming constructs layered on top of VBA. Valuable but ultimately limited by VBA's own expressiveness — you can simulate lambdas, but you can't escape VBA's lack of closures, proper scoping, or modern control flow.

- **Python/COM bridges**: External dependencies, deployment complexity, and inter-process communication overhead.

None of these address the core problem: VBA itself is the bottleneck, and any solution working *within* VBA's syntax is constrained by VBA's grammar.

### ✅ My Solution: ASF as a Language Platform

ASF sidesteps VBA's limitations entirely by implementing a **new language on top of VBA**. Rather than extending VBA's syntax (which is impossible without modifying the runtime), ASF provides a complete alternative syntax that compiles and executes within VBA's runtime environment.

[**🔗 ASF on GitHub**](https://github.com/ws-garcia/ASF)

This means:

- 🔌 **Zero external dependencies** — pure VBA, deployable anywhere VBA runs
- 🔄 **No migration required** — ASF coexists with existing VBA code
- ✨ **Modern features available immediately**: closures, classes with inheritance, modules, functional programming
- ⚡ **VM architecture** enables optimizations impossible in interpreted VBA
- 🛠️ **VS Code tooling** provides a modern development experience on a decades-old platform

### 🧠 Mindset: Bridge Between Eras

Where others see VBA as a dead language to escape from, I see an ecosystem to elevate. Implementing a modern language *inside* VBA rather than *replacing* it reflects how enterprise software actually works: you can't ask organizations to rewrite millions of lines of working code. But you can give them a bridge to modern practices that runs on their existing infrastructure.

I don't abandon users stuck on legacy platforms. I bring the future to them.

---

## 🔁 The Pattern: A Unified Development Philosophy

Across all four cases, the same pattern repeats:

### 1. 🎯 Problem Identification Through Practice

Every project originated from a concrete, practical need — never from abstract curiosity. CSV Interface was needed because Excel couldn't parse non-standard CSV files. Dialect detection was needed for CSV Interface. Expression evaluation was needed for user-defined computations in CSV Interface. ASF was needed because VBA couldn't express the solutions the other problems demanded. Theory follows practice, and each solution reveals the next problem.

### 2. ⚖️ Honest Evaluation of Existing Solutions

In every case, I evaluated existing tools thoroughly and fairly. Excel's import tools were tested against concrete files with measurable expected output. I called stdLambda "fabulous" — and I meant it. I engaged respectfully with sancarn and every community member. But honest evaluation also means identifying **fundamental limitations**, not just surface-level bugs. Excel's non-compliance with RFC-4180, the 50x performance cliff, the positional variable constraint, the structural blindness of character-frequency analysis — these aren't complaints. They're diagnoses.

### 3. 🔀 Reframing the Problem Domain

The decisive move in every case was recognizing that the community was solving the **wrong problem**:

- CSV parsing is not a solved problem — it is a **standards compliance** and **data integrity** challenge
- Dialect detection is not character counting — it is **table structure inference**
- Expression evaluation is not string manipulation — it is **language design**
- VBA modernization is not syntax extension — it is **platform engineering**

Each reframing opened a solution space that was invisible from the original perspective.

### 4. 🏗️ Following the Architecture to Its Conclusion

Once the correct framing was identified, I pursued the solution to its logical end — even when that end was far beyond the original scope. A CSV parser became the foundation for an entire data tools ecosystem. A delimiter guesser became a peer-reviewed research paper. An expression evaluator became a scripting language. A VBA utility became a complete development platform with VS Code tooling. This isn't scope creep — it's the discipline of building the **right** foundation rather than the **quick** patch.

### 5. 📊 Validation Through Rigor

I don't declare my solutions superior — I **measure** them. CSVsniffer's claims rest on F1 scores across multiple datasets compared against multiple tools. ASF's performance claims rest on benchmarks against stdLambda. The Table Uniformity method was submitted to peer review. This commitment to empirical validation transforms personal projects into contributions the broader community can trust and build upon.

### 6. 💪 Resilience Against Dismissal

Perhaps the most defining trait: I pursue ideas the community considers unnecessary, pointless, or excessive — and I persist until the results speak for themselves. The CSV parser dismissed as a "specific use case" became the foundation for everything. The delimiter guesser that was "pointless" is now being discussed for [Python's standard library](https://discuss.python.org/t/rewrite-csv-sniffer/92652). The expression evaluator that wasn't necessary for being a reinvented wheel became a complete language platform. Every dismissal was fuel, not friction.

---

## 📈 Impact Summary

| Project | Origin | Community Verdict | Outcome |
|---|---|---|---|
| 📦 [**VBA CSV Interface**](https://github.com/ws-garcia/VBA-CSV-interface) | Excel fails on non-standard CSV | *"Specific use case"* | Foundation for all subsequent projects; correct where Excel, Power Query, and Power BI fail |
| 🔬 [**CSVsniffer**](https://github.com/ws-garcia/CSVsniffer) | CSV Interface needed delimiter detection | *"Somewhat pointless"* | [Peer-reviewed paper](https://journals.sagepub.com/doi/10.3233/DS-240062) (93.78% F1), [Rust port](https://github.com/jqnatividad/csv-nose) (99.55% on W3C-CSVW), [Python stdlib discussion](https://discuss.python.org/t/rewrite-csv-sniffer/92652) |
| ⚡ **VBA Expressions → [ASF](https://github.com/ws-garcia/ASF)** | CSV Interface needed expression evaluation | *stdLambda suggested as sufficient* | Complete scripting language with VM; outperforms stdLambda |
| 🏗️ [**ASF Platform**](https://github.com/ws-garcia/ASF) | VBA lacks modern features | *twinBASIC / stdLambda suggested* | Full JS-like language in pure VBA, VS Code tooling, module system |

---

## 🎯 Conclusion

The throughline is simple: I treat limitations as specifications for something better. Where the community sees "good enough," I see the gap between what exists and what the problem actually requires. Where others stop at the workaround, I follow the architecture. Where dismissal would discourage most, it sharpens my focus.

The result isn't just a collection of tools — it's a body of work where each piece reinforces the others: Excel's CSV failures demanded CSV Interface, which needed dialect detection, which needed expression evaluation, which needed a proper language, which needed modern tooling. The Table Uniformity method needed rigorous testing, which demanded empirical methodology, which produced peer-reviewed research. Every frustration became a foundation.

**Identify the real problem. Build the real solution. Prove it with real data. Let the results answer every doubt.** 🎯

---

<p align="center">
  <i>If you find this journey interesting, consider ⭐ starring the repositories linked above.</i>
</p>
 

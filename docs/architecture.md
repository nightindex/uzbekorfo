# Architecture

Uzbek Orfo is a Microsoft Word VSTO add-in targeting .NET Framework 4.7.2. It is intentionally maintained as one deployable application because VSTO deployment and Word interop are tightly coupled to the add-in host.

## Repository layout

- `src/UzbekOrfoAddIn/` — VSTO project and product source.
- `eng/` — repository validation scripts.
- `tests/` — tests for logic that can run without Word.
- `docs/` — project documentation.

## Source responsibilities

- `Core/` defines service contracts.
- `Models/` holds application data structures.
- `Services/` implements spelling, grammar, dictionary, export, settings, and workflow behavior.
- `Forms/` and `UI/` contain WinForms presentation and controls.
- `Data/` contains versioned embedded dictionaries and rules.

`ThisAddIn` and the ribbon classes are the VSTO host boundary. `AddInRuntime` is the composition root for runtime services. New domain logic should avoid direct Word or WinForms dependencies where possible so it can be tested in `tests/`.

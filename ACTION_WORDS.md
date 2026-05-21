# Office Agent Action Words

This project accepts Office automation commands from the Next.js assistant, Flask API routes, and the optional global `agent:` keyboard/clipboard listener.

## Global Command Format

Use this form when typing or pasting a command outside the web UI:

```text
agent: <app>: <instruction>
```

Supported app names:

- `excel`
- `word`
- `powerpoint`
- `ppt`

Examples:

```text
agent: excel: create a new workbook
agent: word: create a new document and add heading Project Update
agent: powerpoint: create a presentation with 3 slides about sales
```

## Excel Examples

```text
create a new workbook
open workbook C:/Users/faiza/Desktop/sales.xlsx
save workbook as march_report.xlsx
add a sheet named Summary
rename sheet Sheet1 to Revenue
create a table with 5 rows and 4 columns starting at B2
write 2500 in cell C7
write formula =SUM(B2:B20) in C21
set background color of A1:C1 to yellow
autofit columns A:C
protect sheet Sales with password 1234
```

## Word Examples

```text
create a Word document
open document C:/Users/faiza/Desktop/notes.docx
save document as proposal.docx
add heading Project Update
add paragraph The project is on track
add a table with 3 rows and 5 columns
replace draft with final
set alignment center
set line spacing 1.5
```

## PowerPoint Examples

```text
create a new presentation
open presentation C:/Users/faiza/Desktop/demo.pptx
save presentation as client_pitch.pptx
add a new slide titled Q2 Results
set title on slide 2 to Financial Overview
add bullet point Revenue increased on slide 2
append to body on slide 3 Implementation requires cross-team alignment
add speaker notes to slide 2 This chart shows quarterly growth
insert a table on slide 3 with 4 rows and 3 columns
```

## Compatibility Notes

The action normalizer accepts older LLM response shapes such as:

```json
{"action": "write_cell", "parameters": {"cell": "A1", "value": "Hello"}}
```

and container keys such as `actions`, `commands`, or `steps`. All normalized actions still pass through the central validator before execution.

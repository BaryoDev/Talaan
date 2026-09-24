# Security

## Reporting

**Please do not open a public issue.** Use
[GitHub's private vulnerability reporting](https://github.com/BaryoDev/Talaan/security/advisories/new),
or the contact route in the [BaryoDev policy](https://github.com/BaryoDev/.github/blob/main/SECURITY.md).

Expect an acknowledgement within a few days. This is a small project, so please allow reasonable
time for a fix before disclosing.

## The threat model, in one line

**Talaan parses files it did not create.** The README suggests reading straight from an upload
stream, so assume every input is attacker-controlled and that parsing untrusted bytes is the entire
attack surface. There is no network call, no process launch, no file write and no deserialisation
of arbitrary types anywhere in the library.

## What that means in practice

**Resource exhaustion is the realistic attack.** An `.xlsx` is a zip, and both the zip and the XML
inside it amplify. A malformed or hostile file cannot currently be rejected by size, because there
is no limit anywhere in the read path:

- `XlsxReader.Read` buffers a non-seekable stream into a `MemoryStream` with no cap.
- Each part is handed to `XDocument.Load`, which builds the whole tree in memory.
- The whole grid is materialised into `List<CellValue>` before anything is returned.

Reports that a crafted file makes Talaan allocate far more memory than the file's size are
**in scope and wanted**, ideally with the smallest file that shows it.

**DTD processing is rejected.** Every XML part Talaan reads (workbook.xml, workbook.xml.rels,
sharedStrings.xml, styles.xml, and each worksheet) goes through one reader configured with
`DtdProcessing = Prohibit` and `XmlResolver = null`. A DOCTYPE anywhere, whether it defines
entities for billion-laughs-style expansion or an external entity, is rejected before any of it is
parsed, and surfaces as `InvalidDataException` naming the part (see below). Anything that gets
Talaan to read a file off disk or open a network connection while parsing is still a serious
finding, because nothing in this library is supposed to touch either.

**Zip traversal.** Part paths come from `xl/_rels/workbook.xml.rels` inside the archive, which is
attacker-controlled text. Talaan only ever looks entries up inside the open `ZipArchive` and never
joins those paths onto a filesystem path, so there is no extraction step to traverse out of. If you
find a path that reaches the filesystem, that is a real finding.

## Not vulnerabilities

- **A corrupt file throws.** `InvalidDataException` is the one exception type a bad `.xlsx`
  surfaces as: a file that is not a valid zip, a missing required part, and a part that is not
  well-formed XML or carries a DOCTYPE all throw it. For an XML problem the message names the part
  and `InnerException` is the original `XmlException`. Callers reading uploads need only catch
  `InvalidDataException`. Reports that malformed input throws are expected behaviour; reports that
  malformed input *hangs*, or allocates without bound, are not.
- **Formula results are stale.** Talaan returns the cached `<v>` an authoring tool wrote and never
  evaluates a formula. It cannot execute anything a spreadsheet contains, by construction.
- **Macros are ignored.** `.xlsm` is not a supported extension and macro parts are never read.

## Supported versions

Fixes go to the latest published version on NuGet. There is no long-term support branch.
```

---

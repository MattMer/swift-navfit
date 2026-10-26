# OV-1 High-Level Operational Concept

```mermaid
flowchart LR
  subgraph Command["Command (approved PII storage)"]
    staff["Admin staff and contributors"]
    swift[("Swift NAVFIT workspace<br/>reporting-senior binder")]
    rs["Reporting senior"]
  end
  staff -- "draft, group, review, counsel" --> swift
  rs -- "review, Mark signed" --> swift
  ruad["RUAD roster (CSV)"] -- import --> swift
  older["Earlier Swift workspace"] -- "import (read-only)" --> swift
  pdfin["Structured report PDFs"] -- import --> swift
  swift -- "export / return" --> navfit["NAVFIT98A database"]
  swift -- "print-ready PDFs" --> paper["Official forms for signature"]
  navfit -- "official submission (outside Swift)" --> bol["BUPERS Online"]
  paper -- "signed copies (outside Swift)" --> bol
  swift -- "verified backup" --> backup[("Backup file")]
```

Swift is the command's working binder for performance reports. Administrative
staff and contributors prepare reports, groups and counseling records in it. The
reporting senior reviews them and marks them signed after signing the official
copies. Every exchange with the outside is a file the user chooses. Official
signature and submission stay in the existing Navy process, outside Swift.

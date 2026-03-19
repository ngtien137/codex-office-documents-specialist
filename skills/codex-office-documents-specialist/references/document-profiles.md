# Document Profiles

Use this reference with `scripts/profile_docx.py` when the user asks to analyze a Word document before editing.

## Goals

- identify the likely document type
- identify the likely subject area or topic
- decide whether academic formatting defaults should be offered
- give targeted advice instead of generic cleanup tips

## Recognition Rules

### Academic Thesis Or Dissertation

Signals:

- title or body contains `luan van`, `luan an`, `tieu luan`, `do an`, `bao cao tot nghiep`
- presence of chapter structure such as `Chuong 1`
- front matter markers such as acknowledgements, abstract, table of contents, references, appendices

Action:

- offer `academic-formatting.md` before editing
- if the user accepts, provide a numbered proposed update list and wait for confirmation

### Academic Report

Signals:

- research-language structure such as introduction, literature review, methods, results, discussion, conclusion
- references and appendices present
- less likely to use the exact thesis wording but still clearly academic

Action:

- offer academic formatting when the structure matches a research report
- focus on headings, numbering, captions, and TOC first

### Business Report

Signals:

- KPIs, revenue, cost, market, executive summary, recommendation-heavy sections

Action:

- advise on executive summary, readability, tables, charts, and action-oriented conclusions
- do not force academic rules unless the user explicitly wants them

### Technical Report

Signals:

- architecture, system, API, implementation, testing, deployment, diagrams, appendix-heavy structure

Action:

- advise on problem-method-implementation-results flow
- keep heading hierarchy, captions, and appendix references clean

### Proposal Or Plan

Signals:

- proposal, roadmap, scope, timeline, objectives, implementation plan

Action:

- advise on scope clarity, milestone tables, and summary/recommendation sections

## Topic Advice

### Environment

- pay attention to units, maps, GIS figures, waste or resource-management terminology
- keep figure and table captions explicit

### Education

- separate context, comparison, findings, and recommendations clearly
- make comparative tables easy to scan

### Business And Finance

- tighten headings and summary tables
- verify units, currencies, and time windows

### Technology And Engineering

- keep architecture, method, implementation, and validation distinct
- ensure diagrams and technical tables are captioned clearly

### Health And Biology

- watch scientific names, units, and method/result separation
- keep figure captions self-sufficient

### Law And Policy

- keep section numbering stable
- watch citation order and normative wording consistency

## Required Editing Behavior

If a document is recognized as thesis, dissertation, graduation report, or academic report:

1. ask the user whether they want to update it to the academic formatting profile
2. if the user says yes, inspect and list the exact categories of updates you plan to apply
3. wait for confirmation on the numbered items
4. only then edit, create backup, and log memory

## Analysis Behavior

When the user sends a document only for analysis:

- report the detected document type
- report the likely topic
- summarize the current formatting state
- give tailored advice based on both structure and topic
- mention uncertainty if the detection confidence is low

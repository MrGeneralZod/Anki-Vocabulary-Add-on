# Anki Vocabulary Enricher (starter add-on)

This add-on reads words from a source field and fills other fields using dictionary sources.

## What it can fill

- IPA
- Definition
- Examples
- Synonyms
- Antonyms

## How to use

1. Put this folder in your Anki add-ons directory.
2. Restart Anki.
3. Open **Browse** and select notes you want to update.
4. Click **Tools -> Enrich Vocabulary Fields**.

## Field setup

Your note type should contain:

- a source field (default: `Word`)
- target fields (default: `IPA`, `Definition`, `Examples`, `Synonyms`, `Antonyms`)

You can change names in `config.json`.

## Data sources

- `custom` (default): combines DictionaryAPI + optional Wordnik/Merriam + Datamuse fallback
- `dictionaryapi`: free source, no API key required
- `cambridge`: parses [dictionary.cambridge.org](https://dictionary.cambridge.org) pages (no API key)
- `wordnik`: API key required
- `merriam_webster`: API key required

Pick the source in **Browse -> losev -> Settings**.

## Notes

- Cambridge source may occasionally break if Cambridge changes page markup.
- By default, existing target fields are not overwritten.

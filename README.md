# OutlookToTxt

Simple tool to dump the contents of an Outlook mail folder to plain text files.

## Usage

```
OutlookToTxt --output <target> [--inbox] [--sent] [--folder <foldername> <foldername>]
```

Exports inbox by default. Files will be named based on internal Outlook EntryId.

## Known limitations

- Uses COM interface so requires Outlook client, runs only on Windows
- Loss of fidelity in converting rich text and HTML mails to plain text

## License

MIT
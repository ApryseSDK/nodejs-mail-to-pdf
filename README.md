# pdftron-mail-to-pdf
Convert emails EML, MSG to a PDF with all attachments. The library uses [Mailparser by Nodemailer](https://nodemailer.com/extras/mailparser/) for parsing emails, and PDFTron SDK for conversion from HTML to a PDF, image to PDF and MS Office to PDF. 

## Install
```
npm i
```

## Run
```
npm start
```

## How it works

There are three test emails included with this sample. 

1. We parse the email and get the to, from, subject, and the body of the email as well as the array of attachments.
2. In addition to body's HTML, we add to, from, and subject. You can add additional styling if you wish and retrieve cc.
3. We then convert HTML string to a PDF.
4. If there are any attachments, depending on the content type, we will either convert or just generate the PDF from the buffer of the attachment.
5. Once everything is completed, we merge the PDFs together.

## Project structure

```
files/            - sample emails and resulting converted PDF
  tmp/            - tmp folder needed to write image to disk
```

## API documentation

See [API documentation](https://www.pdftron.com/documentation/web/guides/ui/apis).

## Contributing

See [contributing](./CONTRIBUTING.md).

## License

See [license](./LICENSE).

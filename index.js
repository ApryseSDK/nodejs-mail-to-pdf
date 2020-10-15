const fs = require('fs');
const path = require('path');
const simpleParser = require('mailparser').simpleParser;
const { PDFNet } = require('@pdftron/pdfnet-node');

const OFFICE_MIME_TYPES = [
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
];

const convertEmailToPDF = pathToEmail => {
  fs.readFile(path.resolve(__dirname, pathToEmail), function (err, data) {
    simpleParser(data, {}, (err, parsed) => {
      const html = `
      <div>
        <div>extra div element is needed for padding</div>
        <div><b>from: </b>${parsed.from.html}</div>
        <div><b>to: </b>${parsed.to.html}</div>
        <div><b>subject: </b>${parsed.subject}</div>
      </div><br>${parsed.html}`;

      // create the PDF from the email body
      convertHTMLToPDF(html, 'converted');

      // create PDFs for each of the attachments
      if (parsed.attachments && parsed.attachments.length > 0) {
        parsed.attachments.forEach(async (attachment, i) => {
          const name = `converted_attachment${i}`;
          if (attachment.contentType === 'application/pdf') {
            createPDFAttachment(attachment.content, name);
          } else if (OFFICE_MIME_TYPES.includes(attachment.contentType)) {
            convertFromOffice(attachment.content, name);
          }
        });
      }
    });
  });
};

const convertHTMLToPDF = (html, filename) => {
  const main = async () => {
    try {
      await PDFNet.HTML2PDF.setModulePath(
        path.resolve(__dirname, './node_modules/@pdftron/pdfnet-node/lib/'),
      );
      const outputPath = path.resolve(__dirname, `./files/${filename}.pdf`);
      const html2pdf = await PDFNet.HTML2PDF.create();
      const pdfdoc = await PDFNet.PDFDoc.create();
      await html2pdf.insertFromHtmlString(html);
      await html2pdf.convert(pdfdoc);
      await pdfdoc.save(outputPath, PDFNet.SDFDoc.SaveOptions.e_linearized);
    } catch (err) {
      console.log(err);
    }
  };

  PDFNetEndpoint(main);
};

const createPDFAttachment = (buffer, filename) => {
  const main = async () => {
    try {
      const outputPath = path.resolve(__dirname, `./files/${filename}.pdf`);
      const pdfdoc = await PDFNet.PDFDoc.createFromBuffer(buffer);
      await pdfdoc.save(outputPath, PDFNet.SDFDoc.SaveOptions.e_linearized);
    } catch (err) {
      console.log(err);
    }
  };

  PDFNetEndpoint(main);
};

const convertFromOffice = (buffer, filename) => {
    const main = async () => {
      try {
        const outputPath = path.resolve(__dirname, `./files/${filename}.pdf`);
        const pdfdoc = await PDFNet.PDFDoc.createFromBuffer(buffer);
        await pdfdoc.save(outputPath, PDFNet.SDFDoc.SaveOptions.e_linearized);
      } catch (err) {
        console.log(err);
      }
    };
  
    PDFNetEndpoint(main);
  };

const PDFNetEndpoint = main => {
  PDFNet.runWithCleanup(main) // you can add the key to PDFNet.runWithCleanup(main, process.env.PDFTRONKEY)
    .then(() => {
      PDFNet.shutdown();
    })
    .catch(err => {
      console.log(err);
    });
};

convertEmailToPDF('./files/test2attachment.eml');

const express      = require('express'),
      multer       = require('multer'),
      upload       = multer({storage: multer.memoryStorage()}),
      MemoryStream = require('memorystream'),
      moment       = require('moment'),
      excellent    = require('excellent'),
      parse        = require('csv').parse,
      fs           = require('fs'),
      https        = require('https'),
      request      = require('request'),
      cheerio      = require('cheerio'),
      app          = express();

const zip  = new require('node-zip')(),
      port = process.env.PORT || 3060,
      date = moment();

app.get('/', (req, res) => {
    res.send(`<!DOCTYPE html>
<html>
    <body>
        <form enctype="multipart/form-data" method="post">
            <input type="text" name="name" placeholder="Name" />
            <input type="text" name="department" placeholder="Department" />
            <input type="file" name="file" accept=".csv" />
            
            <input type="submit" value="Generate" />
        </form>
    </body>
</html>

`);
});

app.post('/', upload.single('file'), (req, res) => {
    console.log(`Converting ${req.file.originalname} to the SCTR Format.`);

    parse(req.file.buffer, {delimiter: ','}, (err, output) => {
        if (err) {
            res.send("Failed to parse csv. " + err);

            return;
        }

        generateReport(output, req.body.name, req.body.department)
            .catch(err => {
                res.send("Failed to generate report. " + err);
            })
            .then(doc => {
                try {
                    zip.file('report.xlsx', doc.file);

                    res.writeHead(200, {
                        'Content-Type':        'application/zip',
                        'Content-disposition': `attachment; filename=report_${req.body.name.replace(/ /g, '_')}.zip`
                    });

                    let data = zip.generate({base64: false, compression: 'DEFLATE'});
                    res.end(data, 'binary');

                    console.log("Created report!");
                } catch (e) {
                    console.log(e);
                    res.send("Failed to zip report. " + e);
                }
            });
    });
});

app.listen(port, () => {
    console.log("Listening on http://localhost:" + port);
});

function generateReport(csv, name, department) {
    return new Promise(resolve => {
        let sheetName = 'Report - ' + date.format("MMM Do YYYY"),
            report    = {
                sheets: {
                    [sheetName]: {
                        image: {image: fs.readFileSync(__dirname + '/logo.png'), filename: 'logo.png'},
                        rows:  [
                            {cells: ['']},
                            {cells: ['']},
                            {cells: ['']},
                            {cells: ['']},
                            {cells: ['']},
                            {cells: ['']},
                            {cells: ['']}, // First eight rows are blank.
                            {
                                cells: [
                                    {value: 'Name: ', style: 'bold'},
                                    {value: name, style: 'underline'},
                                ]
                            },
                            {
                                cells: [
                                    {value: 'Dept: ', style: 'bold'},
                                    {value: department, style: 'underline'},
                                ]
                            },
                            {
                                cells: [
                                    {value: 'Date: ', style: 'bold'},
                                    {value: date.format("MMM Do YYYY"), style: 'underline'},
                                ]
                            },
                            {
                                cells: [
                                    {value: 'Reason: ', style: 'bold'},
                                    {value: 'Expense Report', style: 'underline'},
                                ]
                            },
                            {cells: ['']}, // Blank line
                            {
                                cells: [
                                    {value: "DATE", style: 'header'},
                                    {value: "MERCHANT", style: 'header'},
                                    {value: "CHARGE", style: 'header'},
                                    {value: "NOTES", style: 'header'}
                                ]
                            }
                        ]
                    }
                },
                styles: {
                    borders: [
                        {
                            label:  'bottom',
                            bottom: {style: 'thin', color: 'Black'}
                        }
                    ],
                    fonts:   [{label: 'bold', bold: true}],
                    fills:   [{label: 'yellow', type: 'pattern', color: 'Yellow'}],

                    cellStyles: [
                        {label: 'bold', font: 'bold'},
                        {label: 'underline', borderBottom: 'black'},
                        {label: 'header', fill: 'yellow', font: 'bold'}
                    ]
                }
            };

        csv.shift();
        csv.forEach(row => {
            let notes = '';
            for (let i of [5, 6]) {
                if (row[i]) {
                    if (notes.length > 0) {
                        notes += ', ';
                    }

                    notes += row[i];
                }
            }

            report.sheets[sheetName].rows.push({
                cells: [
                    moment(new Date(row[0])).format('MM/DD/YYYY'),
                    row[1],
                    row[2],
                    notes
                ]
            });
        });

        report.sheets[sheetName].rows.push({cells: ['']});
        report.sheets[sheetName].rows.push({cells: ['']});
        report.sheets[sheetName].rows.push({
            cells: [
                '',
                'Total: ',
                '$' + csv.map(i => i[2]).reduce((a, b) => a + b)
            ]
        });

        Promise.all(DownloadFiles(csv))
            .catch(e => console.log(e))
            .then(() => {
                console.log("Finished Downloading receipts");
                resolve(excellent.create(report))
            });
    });
}

function DownloadFiles(csv) {
    return csv.map(row => {
        return new Promise(resolve => {
            request(row[10], (err, response, html) => {
                if (err) {
                    console.log("One of the links for the reciepts is bad.");
                    process.exit(1);
                }

                let $      = cheerio.load(html, {xmlMode: false}),
                    script = $($('script')).text();

                function findTextAndReturnRemainder(target, variable) {
                    let chopFront = target.substring(target.search(variable) + variable.length, target.length),
                        result    = chopFront.substring(0, chopFront.search(";"));


                    console.log(result);
                    try {
                        return JSON.parse(result);
                    } catch (e) {
                        throw new Error(e, result, target);
                    }
                }

                let transaction = findTextAndReturnRemainder(script, "var transaction =");

                let url = 'https://s3.amazonaws.com/receipts.expensify.com/' + transaction.receiptFilename;
                console.log('downloading ' + url);
                let fileName = transaction.merchant.replace(/ /g, '_') + '.pdf';

                let memStream = new MemoryStream(null, {readable: false});
                https.get(url, res => {
                    res.pipe(memStream);
                    res.on('end', () => {
                        zip.folder('receipts').file(fileName, memStream.toString());

                        resolve();
                    });
                    res.on('error', err => {
                        console.log(err)
                    })
                }).on('error', err => {
                    console.log(err)
                });
            })
        })
    });
}

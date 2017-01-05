const express    = require('express'),
      multer     = require('multer'),
      upload     = multer({storage: multer.memoryStorage()}),
      moment     = require('moment'),
      excellent  = require('excellent'),
      parse      = require('csv').parse,
      fs         = require('fs'),
      mime       = require('mime-types'),
      https      = require('https'),
      request    = require('request'),
      cheerio    = require('cheerio'),
      app        = express();

const zip  = new require('node-zip')(),
      port = process.env.PORT || 3060,
      date = moment();

app.get('/', (req, res) => {
    res.send(`<!DOCTYPE html>
<html>
    <head>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bulma/0.3.0/css/bulma.min.css">
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css">
    </head>
    <body>
        <section class="hero is-primary is-fullheight">
            <div class="hero-head">
                <header class="nav">
                    <div class="nav-center">
                        <h2 class="title nav-item is-2">
                            SCTR Expense Report
                        </a>
                    </div>
                </header>
            </div>
            <div class="hero-body">
                <div class="container has-text-centered">
                    <form enctype="multipart/form-data" method="post">
                        <div class="control is-grouped">
                            <div class="control is-horizontal is-expanded">
                                <div class="control-label">
                                    <label class="label">Name</label>
                                </div>
                                <div class="control">
                                    <input type="text" name="name" placeholder="Name" class="input"/>
                                </div>
                            </div>
                            <div class="control is-horizontal is-expanded">
                                <div class="control-label">
                                    <label class="label">Department</label>
                                </div>
                                <div class="control">
                                    <input type="text" name="department" placeholder="Department" class="input"/>
                                </div>
                            </div>
                        </div>
                        <div class="control is-grouped">
                            <div class="control is-horizontal is-expanded">
                                <div class="control-label">
                                    <label class="label">File</label>
                                </div>
                                <div class="control">
                                    <input type="file" name="file" class="file" accept=".csv" />
                                </div>
                            </div>
                            <div class="control is-horizontal is-expanded">
                                <input class="button is-info is-fullwidth" type="submit" value="Generate" />
                            </div>
                        </div>
                    </form>
                </div>
            </div>
        </section>
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
                    zip.file(`report_${req.body.name.replace(/ /g, '_')}.xlsx`, doc.file);

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
                    '$' + row[2].replace(/['"]/g, ''),
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
                '$' + csv.map(i => i[2]).reduce((a, b) => parseInt(a) + parseInt(b))
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
            if (!row[10]) {
                resolve();
                return;
            }
            request(row[10], (err, response, html) => {
                if (err) {
                    console.log("One of the links for the receipts is bad.");
                    resolve();
                    return;
                }

                let $      = cheerio.load(html, {xmlMode: false}),
                    script = $($('script')).text();

                function findTextAndReturnRemainder(target, variable) {
                    let chopFront = target.substring(target.search(variable) + variable.length, target.length),
                        result    = chopFront.substring(0, chopFront.search(";"));

                    try {
                        return JSON.parse(result);
                    } catch (e) {
                        throw new Error(e, result, target);
                    }
                }

                let transaction = findTextAndReturnRemainder(script, "var transaction =");

                let url      = 'https://s3.amazonaws.com/receipts.expensify.com/' + transaction.receiptFilename;
                let fileName = getFilename(transaction, response.headers['content-type']);

                if (!transaction.receiptFilename) {
                    zip.folder('receipts').file(fileName + '.txt', row[10]);
                    resolve();
                    return;
                }

                let file = fs.createWriteStream("/tmp/" + fileName);
                request({uri: url}).pipe(file).on('close', () => {
                    file.end();
                    try {
                        fs.readFile("/tmp/" + fileName, (err, f) => {
                            if (err) {
                                console.log("One of the links for the receipts is bad.");
                                resolve();
                                return;
                            }

                            try {
                                zip.folder('receipts').file(fileName, f);
                            } catch (e) {
                                console.log(e);
                            }

                            resolve();
                        })
                    } catch (e) {
                        console.log(e)
                    }
                });
            })
        })
    });
}

function getFilename(transaction) {
    const extension = /(?:\.([^.]+))?$/.exec(transaction.receiptFilename)[1],
          name      = (transaction.modifiedMerchant || transaction.merchant).replace(/ /g, '_') + "_" + transaction.created;

    return name + (extension ? '.' + extension : '');
}

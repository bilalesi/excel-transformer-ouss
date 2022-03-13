const _ = require('lodash');
const writeXlsxFile = require('write-excel-file/node');
const readXlsxFile = require('read-excel-file/node');
const chalk = require('chalk');
const cliProgress = require('cli-progress');
const path = require("path");
const { nanoid } = require('nanoid');


const treat_data_and_export_xlsx = async () => {
    console.log(chalk.green('Reading data from file...'));
    const barProgress = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);
    barProgress.start(100, 0);
    let file = path.join(__dirname, 'job.xlsx');
    try {
        console.log(chalk.green('Data readed!'));
        const rows = await readXlsxFile(file);
        console.log(chalk.green('Treating data...'));
        barProgress.update(20);
        let real_data = rows.length > 0 && rows.map(item => ([
            [1, "GTIN", item[0]],
            [3, "IMAGE PRODUIT", "" ],
            [4, "DESIGNATION DE VENTE", item[3]],
            [5, "RAISON SOCIAL", item[4]],
            [6, "MARQUE DEPOSE", item[5]],
            [7, "ADRESSE IMPORTATEUR", item[6]],
            [8, "PAYS D'ORIGINE OU PAYS D'IMPORTATION", item[7]],
            [9, "Marque de conformité liée à la sécurité", item[8]],
            [10, "Référentiel de pré-licence pour les produits concernés", item[9]],
            [11, "Quantité nette exprimée dans le système métrique international", item[10]],
            [12, "Précautions prises dans le domaine de la sécurité", item[11]],
            [13, "Composants du produit et conditions de stockage", item[12]],
            [14, "Toutes les autres informations utiles peuvent également être ajoutées", item[13]],
            ["", "", ""], [""," ", ""],
        ]
        )).map(item => {
            barProgress.increment(1);
            return item.map((cell, index) => {
                if (index === item.length - 1 || index === item.length - 2) {
                    return cell.map((sub, ind) => ({ value: sub ? sub : "", align: "left", alignVertical: 'center' }))
                }
                return cell.map((sub, ind) => ({ value: sub ? sub : "", align: "left", alignVertical: 'center', borderStyle: "thin", borderColor: "#333", }))
            })
        })
        let file_name = `${nanoid(2)}_export.xlsx`;
        await writeXlsxFile(_.flatten(real_data), {
          columns: [
            { column: 'Number', key: 'value', width: 20 },
            { column: 'Title', key: 'value', width: 80 },
            { column: 'Data', key: 'value', width: 60 },
          ],
          fontSize: 14,
          sheet: `sheet-${nanoid(3)}`,
          filePath: `${__dirname}/${file_name}`,
        })
        barProgress.update(100);
        barProgress.stop();
        console.log(chalk.green('Data treated!'));
        console.log(chalk.green('File: '), file_name);
    } catch (error) {
        console.log(chalk.red("Error",), error);
    }
}

treat_data_and_export_xlsx();
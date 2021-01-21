const express = require('express');
const { promisify } = require('util');
const app = express();


app.use(express.static('./pages/public'))

const nunjucks = require('nunjucks');
nunjucks.configure('./pages/', {
    express: app
});



const docId = '1Mjr-kxuzJCbh53pXKdUXmLctKnaLZF5KINg9BxSXbko';



const { GoogleSpreadsheet } = require('google-spreadsheet');
const { ALPN_ENABLED } = require('constants');

var datas = [];
let totalfaltas = 0 ;




app.get('/', async (request, response) => {

    await  carregar();
   

    return response.render("index.html",{datas});
})
 


app.get('/tarefa', async (request, response) => {

    const doc = new GoogleSpreadsheet(docId);

    // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth({
        client_email: "tarefa-google@quickstart-1610739987253.iam.gserviceaccount.com",
        private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDEcaJLmv99DWVh\nSXiNR6biFwrSG1fQ2nqKNjy2DRdLppVCWU2L3sHWD0AJa5IApHpgBCrKUWKXUcOO\nQIjS9aMHXo5LJiChnPGtZntCXN3BKOUFIe8+mxhUGBzNGwSwlTOzbcSTcUhLhfmU\nUkEL1rDF+UkB23j+UaH3UUZewN1dFT3EkEa8GFRs8b/Wupw8S44CW+bTLGolccEA\nlEYsq9OZnLOxcmZIiWOF9afQggvtDgG7NcGhOz3u5GJw3QR1/FG6OmPTDdyH8B1K\ncnq4kzBVRwhaSZnxhaGKvxNM19vBt14Z5ZGZ2VOmrXD37C6p+nmHMxRP7hkOvN3i\nQXBoAbRxAgMBAAECggEABgNManAGHffJAJ9VF034J7d411GK8JOfaJecaB4idmhU\n7UD6hKt+12SEG0W1pFtke4flH2g6UlNoXvROu9ZU9SbJyDcUjJ3XL+2RHEjnaMAt\nsmiFgC8TIY/TYdvP2u/WM0nK2JCBG/6v0wBpiUk7A/RLbckf/PjWslFEjCXvIKg2\nDwI6peqdsNdGr6lHmVoA33foZSYExY1zNuTqr5RsGWsBDGflGapb1VcUV9QE40yX\nRj3LgsCOPsROGkvbGrhnisw8E+Ea0rm+kyRpa5lCdWhCKUE8ue7vjpEDmvLlir6P\nX6BqvPFO+apf58A/i/nRRgJbFC9VY5kr6Z+OuD4W+QKBgQDjSAUwT8/J8A4bfeB+\nuSD6xzDuRokBobE7RSSTcc1G1IAH8RbJB8D6Q/3jmTl3uI9NnEz6S6aHVM5rJz1l\nO2QDJwUSP1qMLenF72mUF0iBxMqeEewsXhMLZ+ZYkCZCpEk1VcrUEDaj9QOWcmK6\nIvZ+rbic7TnbKoj0CnvfkEJ0MwKBgQDdRBmNV/VifEFfkcYL0hWeBLQMk/zAHV/y\nKa5Ivo5ygFArMEkLkgQzL7rOTzq0c5q0x7zo9JVaCc7mXYl+Oumxc9fKiyB26eeM\nE3F5pMc6mlk/dyZMwbApuZ4A7mZ/IHrn6aaP7rPMJ+uA6xnAjwJJ/ojgN8Wm7dSE\nqyxhpmIwywKBgHDzx/Bcmc2oCbrL8hfIdYVsHPst/sTa0LO+BxFnyzbaQM6xmDtM\nKTG3PKQx8Ad5p25QsUjq89Xp5bQHClIXE/slFzYcWim0X6vI8dVxRM2JOZEZIyBh\nmGFgv29gJEOWVfO1sVl2vVD6YVARhNMwsQP/3fHPS6OKHgn6c9mFXiFVAoGAe8PF\nzyvuFAKQxpZRgvcmJFdZJtf4PrWvn1L1K7d7Ekz3itDdat1oAAGoqhHjMmCfnpNC\n9cMpb02hL3YOnE7zvNChWafsptc7Lz0I8hPbZMpFNZy+DZ0hnpU27iprppxSYzps\ncoIAjCegMWJP60eS7jSz90b7Bd5uSy88Cfr5XXUCgYAzyuIJtk8ZDkn2Jelvqodg\n1TpZ55JQU0v+KEx12xyTuscp1h9zuBkh03FOCat/wgKHucCmAwASXCFUJ99R2b82\nYyEvxwFwNlbHsfdb7yYy1Dy1X7HWWtioONRoaoquPQOLGA1sR4ikCp6UgfFCGNSa\nP43mRUw+9u9lvgPn4R1Umw==\n-----END PRIVATE KEY-----\n",
    });

    await doc.loadInfo(); // loads document properties and worksheets
    console.log(doc.title);
    await doc.updateProperties({ title: 'renamed doc' });
    const sheet = doc.sheetsByIndex[0]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
    await sheet.loadCells('A2:H27'); // loads a range of cells

    let total = 26;
    totalfaltas = Number(sheet.getCell(1, 0).value.split(":")[1]);
    for (i = 3; i <= total; i++) {


        let cellSituacao = sheet.getCell(i, 6);
        let cellFinal = sheet.getCell(i, 7);
        let faltas = sheet.getCell(i, 2).value;
        let p1 = sheet.getCell(i, 3).value;
        let p2 = sheet.getCell(i, 4).value;
        let p3 = sheet.getCell(i, 5).value;



        cellSituacao.value = situacao(p1, p2, p3, faltas)
        cellFinal.value = notaFinal(p1, p2, p3, faltas,);
        
        await sheet.saveUpdatedCells();
    }

   

    await carregar()

   
        return response.render("index.html",{datas});
   

    

});

async function  carregar(){
    const doc = new GoogleSpreadsheet(docId);

    // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
    await doc.useServiceAccountAuth({
        client_email: "tarefa-google@quickstart-1610739987253.iam.gserviceaccount.com",
        private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDEcaJLmv99DWVh\nSXiNR6biFwrSG1fQ2nqKNjy2DRdLppVCWU2L3sHWD0AJa5IApHpgBCrKUWKXUcOO\nQIjS9aMHXo5LJiChnPGtZntCXN3BKOUFIe8+mxhUGBzNGwSwlTOzbcSTcUhLhfmU\nUkEL1rDF+UkB23j+UaH3UUZewN1dFT3EkEa8GFRs8b/Wupw8S44CW+bTLGolccEA\nlEYsq9OZnLOxcmZIiWOF9afQggvtDgG7NcGhOz3u5GJw3QR1/FG6OmPTDdyH8B1K\ncnq4kzBVRwhaSZnxhaGKvxNM19vBt14Z5ZGZ2VOmrXD37C6p+nmHMxRP7hkOvN3i\nQXBoAbRxAgMBAAECggEABgNManAGHffJAJ9VF034J7d411GK8JOfaJecaB4idmhU\n7UD6hKt+12SEG0W1pFtke4flH2g6UlNoXvROu9ZU9SbJyDcUjJ3XL+2RHEjnaMAt\nsmiFgC8TIY/TYdvP2u/WM0nK2JCBG/6v0wBpiUk7A/RLbckf/PjWslFEjCXvIKg2\nDwI6peqdsNdGr6lHmVoA33foZSYExY1zNuTqr5RsGWsBDGflGapb1VcUV9QE40yX\nRj3LgsCOPsROGkvbGrhnisw8E+Ea0rm+kyRpa5lCdWhCKUE8ue7vjpEDmvLlir6P\nX6BqvPFO+apf58A/i/nRRgJbFC9VY5kr6Z+OuD4W+QKBgQDjSAUwT8/J8A4bfeB+\nuSD6xzDuRokBobE7RSSTcc1G1IAH8RbJB8D6Q/3jmTl3uI9NnEz6S6aHVM5rJz1l\nO2QDJwUSP1qMLenF72mUF0iBxMqeEewsXhMLZ+ZYkCZCpEk1VcrUEDaj9QOWcmK6\nIvZ+rbic7TnbKoj0CnvfkEJ0MwKBgQDdRBmNV/VifEFfkcYL0hWeBLQMk/zAHV/y\nKa5Ivo5ygFArMEkLkgQzL7rOTzq0c5q0x7zo9JVaCc7mXYl+Oumxc9fKiyB26eeM\nE3F5pMc6mlk/dyZMwbApuZ4A7mZ/IHrn6aaP7rPMJ+uA6xnAjwJJ/ojgN8Wm7dSE\nqyxhpmIwywKBgHDzx/Bcmc2oCbrL8hfIdYVsHPst/sTa0LO+BxFnyzbaQM6xmDtM\nKTG3PKQx8Ad5p25QsUjq89Xp5bQHClIXE/slFzYcWim0X6vI8dVxRM2JOZEZIyBh\nmGFgv29gJEOWVfO1sVl2vVD6YVARhNMwsQP/3fHPS6OKHgn6c9mFXiFVAoGAe8PF\nzyvuFAKQxpZRgvcmJFdZJtf4PrWvn1L1K7d7Ekz3itDdat1oAAGoqhHjMmCfnpNC\n9cMpb02hL3YOnE7zvNChWafsptc7Lz0I8hPbZMpFNZy+DZ0hnpU27iprppxSYzps\ncoIAjCegMWJP60eS7jSz90b7Bd5uSy88Cfr5XXUCgYAzyuIJtk8ZDkn2Jelvqodg\n1TpZ55JQU0v+KEx12xyTuscp1h9zuBkh03FOCat/wgKHucCmAwASXCFUJ99R2b82\nYyEvxwFwNlbHsfdb7yYy1Dy1X7HWWtioONRoaoquPQOLGA1sR4ikCp6UgfFCGNSa\nP43mRUw+9u9lvgPn4R1Umw==\n-----END PRIVATE KEY-----\n",
    });

    await doc.loadInfo(); // loads document properties and worksheets
    console.log(doc.title);
    await doc.updateProperties({ title: 'renamed doc' });
    const sheet = doc.sheetsByIndex[0]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
    await sheet.loadCells('A2:H27'); // loads a range of cells
  

    const rows = await sheet.getRows(); // can pass in { limit, offset }
    const totalLinha = rows.length 

    datas = [];

     for( i = 3 ; i <= totalLinha ; i ++ ){
     
        datas.push({"matricula":sheet.getCell(i, 0).value,
        "nome":sheet.getCell(i, 1).value,
        "faltas":sheet.getCell(i, 2).value,
        "p1":sheet.getCell(i, 3).value,
        "p2":sheet.getCell(i, 4).value,  
        "p3":sheet.getCell(i, 5).value,
        "situacao":sheet.getCell(i, 6).value,
        "nota_final":sheet.getCell(i, 7).value});
     }
}


  function situacao(p1, p2, p3, faltas) {
   
    if (faltas > (totalfaltas / 4)) {
        return "Reprovado por falta";
    }

    let media = (p1 + p2 + p3) / 3;

    if (media >= 70) {
        return "Aprovado"
    } else if (media <= 50 && media < 70) {
        return "Prova final"
    } else {
        return "Reprovado por Nota"
    }

}

function notaFinal (p1, p2, p3, faltas){
    
    if (faltas > (totalfaltas / 4)) {
        return '';
    }

    let media = (p1 + p2 + p3) / 3;

    if(media >= 50 && media < 70 ){
       let  nota =  (70 - media );
        return  Math.round( media + ( nota * 2)) ;
    } else if (media >= 70){
        return 0;
    }

}


 


app.listen(3333, () => {
    console.log('connected server');
});
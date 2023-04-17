process.setMaxListeners(0)
const Excel = require('exceljs');
const puppeteer = require('puppeteer');
const chr = require("cheerio");
const fs = require("fs");
const readlineSync = require('readline-sync');
const path = require('path');

const LAUNCH_PUPPETEER_OPTS = {
    headless: true,
    args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--disable-gpu',
        '--window-size=1920x1080'
    ]
};

const PAGE_PUPPETEER_OPTS = {
    networkIdle2Timeout: 5000,
    waitUntil: 'networkidle2',
    timeout: 3000000
};

const getHTML = async (url) => {
    const browser = await puppeteer.launch(LAUNCH_PUPPETEER_OPTS);
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(0);
    await page.goto(url, PAGE_PUPPETEER_OPTS);
    const content = await page.content();
    await page.close(); // MB UBRAT
    await browser.close();
    return chr.load(content);
};

const parseSellers = async () => {
    try {
        let num = 13387012;
        let posts = [];
        let post = {};
        let $1;
        let arrOfLinks = [];

        for (let i = 1; i <= 1000; i++) {
            $ = await getHTML(`https://kaspi.kz/shop/info/merchant/${num}/address-tab/`);
            if ($(".merchant-profile__name").text().length == 0) {
                num--;
                continue;
            }
            console.log(`https://kaspi.kz/shop/info/merchant/${num}/address-tab/`);
            arrOfLinks.push(`https://kaspi.kz/shop/info/merchant/${num}/address-tab/`);
            num--;
        }

        console.log(arrOfLinks);
        for (const element of arrOfLinks) {
            console.log(element)
            post = {};
            try {
                $1 = await getHTML(element);
                const name = $1(".merchant-profile__name").text();
                const phoneNumber = $1(".merchant-profile__contact-text").text();
                const link = element;

                post.name = name;
                post.phoneNumber = phoneNumber;
                post.link = link;

                posts.push(post);
            } catch (error) {
                console.log(error, ": " + element)
            }
        }
        await createExcel(posts)
    } catch (error) {
        console.log(error);
    }
}

const createExcel = async (posts) => {
    if(fs.existsSync(path.join(__dirname, 'sellers/sellers.xlsx'))){
        fs.unlinkSync(path.join(__dirname, 'sellers/sellers.xlsx'), err => {
            if (err) throw err;
            console.log('Файл успешно удалён');
        });
    }
    try {
        let lengthCell = 0;
        let lengthCellPhone = 0;
        workbook = new Excel.Workbook();
        worksheet = workbook.addWorksheet('Продавцы');
        worksheet.getRow(1).values = ['Название продавца', 'Номер телефона', 'Ссылка на магазин'];
        worksheet.columns = [
            { key: 'name' },
            { key: 'phoneNumber' },
            { key: 'link' }
        ];
        posts.forEach(element => {
            if (lengthCell < element.name.length) lengthCell = element.name.length;
            if (lengthCellPhone < element.phoneNumber.length) lengthCellPhone = element.phoneNumber.length;
        });
        worksheet.columns.forEach((column, i) => {
            if (i == 0) column.width = lengthCell + 5;
            else if (i == 1) column.width = lengthCellPhone + 5
            else column.width = worksheet.getRow(1).values[i].length + 5;
        })
        worksheet.getRow(1).font = { bold: true };
        posts.forEach((e, index) => {
            worksheet.addRow({ ...e });
        })
        await workbook.xlsx.writeFile(path.join(__dirname, 'sellers/sellers.xlsx'))
    } catch (error) {
        console.log(error)
    }
}

(async function () {
    parseSellers().then(() => {
        readlineSync.question('Sellers have been parsed. Look into sellers folder\nPress enter to exit');
    })
}())

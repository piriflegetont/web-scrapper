"use strict";

const fs = require('fs');
const excel = require('excel4node');
const math = require('mathjs');
const cheerio = require('cheerio');
const rp = require('request-promise');
const https = require('https');
const path = require('path');
const moment = require('moment');
const Progress = require('./progress');

const FOLDER = 'parsingAt_' + moment().format('DD.MM.YYYY_HH.mm');
const REQ_IN_TIME = 200;
const TIMES_TO_TRY = 10;

async function makeRequest(url, parse, i) {
	try {
		let options = {
			method: 'GET', uri: url, resolveWithFullResponse: true
		};
		let responce = await rp(options);
		if (responce.statusCode === 200) {
			let html = responce.body;
			let $ = cheerio.load(html);
			return await parse($, url);
		} else {
			throw new Error('Responce code is not 200');
		}
	} catch (err) {
		if (typeof i === 'undefined') {
			i = 0;
		}
		console.log('again ', i);
		if (i >= TIMES_TO_TRY) {
			throw err;
		}
		return makeRequest(url, parse, i + 1)
	}
}

function saveAsXLSX(objects) {
	try {
		let workbook = new excel.Workbook();
		let worksheet = workbook.addWorksheet('Sheet 1');
		let style = workbook.createStyle({
			font: {
				color: '#000000', size: 9
			}, numberFormat: '$#,##0.00; ($#,##0.00); -'
		});
		worksheet.cell(1, 1).string('наименование').style(style);
		worksheet.cell(1, 2).string('артикул').style(style);
		worksheet.cell(1, 3).string('объём').style(style);
		worksheet.cell(1, 4).string('описание').style(style);
		worksheet.cell(1, 5).string('мнение специалиста').style(style);
		worksheet.cell(1, 6).string('Страна производства').style(style);
		worksheet.cell(1, 7).string('Производители масел').style(style);
		worksheet.cell(1, 8).string('По производителю автомобиля').style(style);
		worksheet.cell(1, 9).string('Область применения').style(style);
		worksheet.cell(1, 10).string('Тип продукта').style(style);
		worksheet.cell(1, 11).string('Вязкость по SAE').style(style);
		worksheet.cell(1, 12).string('API').style(style);
		worksheet.cell(1, 13).string('ACEA').style(style);
		worksheet.cell(1, 14).string('Спецификации производителей автомобилей').style(style);
		worksheet.cell(1, 15).string('Допуски производителей автомобилей').style(style);
		worksheet.cell(1, 16).string('Тип двигателя').style(style);
		worksheet.cell(1, 17).string('Тип топлива').style(style);
		worksheet.cell(1, 18).string('Срок годности, мес').style(style);
		worksheet.cell(1, 19).string('Внешний вид масла, смазки').style(style);
		worksheet.cell(1, 20).string('Цвет продукта').style(style);
		worksheet.cell(1, 21).string('Испаряемость по НОАК').style(style);
		worksheet.cell(1, 22).string('Индекс вязкости').style(style);
		worksheet.cell(1, 23).string('Вязкость кинематическая при 40°С').style(style);
		worksheet.cell(1, 24).string('Вязкость кинематическая при 100°С').style(style);
		worksheet.cell(1, 25).string('Вязкость кажущаяся (динамическая), определяемая на имитаторе холодной прокрутки (CCS) при -30С').style(style);
		worksheet.cell(1, 26).string('Плотность при +15 цельсия').style(style);
		worksheet.cell(1, 27).string('Температура застывания').style(style);
		worksheet.cell(1, 28).string('Температура вспышки').style(style);
		worksheet.cell(1, 29).string('Температура потери текучести').style(style);
		worksheet.cell(1, 30).string('Общее щелочное число (TBN)').style(style);
		worksheet.cell(1, 31).string('Общее кислотное число (TAN)').style(style);
		worksheet.cell(1, 32).string('Массовая доля серы').style(style);
		worksheet.cell(1, 33).string('Зола сульфатная').style(style);
		worksheet.cell(1, 34).string('Содержание цинка').style(style);
		worksheet.cell(1, 35).string('Содержание фосфора').style(style);
		worksheet.cell(1, 36).string('Содержание бора').style(style);
		worksheet.cell(1, 37).string('Содержание магния').style(style);
		worksheet.cell(1, 38).string('Содержание кальция').style(style);
		worksheet.cell(1, 39).string('Содержание натрия').style(style);
		for (let i = 0; i < objects.length; i++) {
			try {
				worksheet.cell(i + 2, 1).string(objects[i].naimen).style(style);
				worksheet.cell(i + 2, 2).string(objects[i].articul).style(style);
				worksheet.cell(i + 2, 3).string(objects[i].volume).style(style);
				worksheet.cell(i + 2, 4).string(objects[i].descr).style(style);
				worksheet.cell(i + 2, 5).string(objects[i].specialist).style(style);
				worksheet.cell(i + 2, 6).string(objects[i].table1['Страна производства']).style(style);
				worksheet.cell(i + 2, 7).string(objects[i].table1['Производители масел']).style(style);
				worksheet.cell(i + 2, 8).string(objects[i].table1['По производителю автомобиля']).style(style);
				worksheet.cell(i + 2, 9).string(objects[i].table1['Область применения']).style(style);
				worksheet.cell(i + 2, 10).string(objects[i].table1['Тип продукта']).style(style);
				worksheet.cell(i + 2, 11).string(objects[i].table1['Вязкость по SAE']).style(style);
				worksheet.cell(i + 2, 12).string(objects[i].table1['API']).style(style);
				worksheet.cell(i + 2, 13).string(objects[i].table1['ACEA']).style(style);
				worksheet.cell(i + 2, 14).string(objects[i].table1['Спецификации производителей автомобилей']).style(style);
				worksheet.cell(i + 2, 15).string(objects[i].table1['Допуски производителей автомобилей']).style(style);
				worksheet.cell(i + 2, 16).string(objects[i].table1['Тип двигателя']).style(style);
				worksheet.cell(i + 2, 17).string(objects[i].table1['Тип топлива']).style(style);
				worksheet.cell(i + 2, 18).string(objects[i].table1['Срок годности, мес']).style(style);
				worksheet.cell(i + 2, 18).string(objects[i].table2['Внешний вид масла, смазки']).style(style);
				worksheet.cell(i + 2, 19).string(objects[i].table2['Цвет продукта']).style(style);
				worksheet.cell(i + 2, 20).string(objects[i].table2['Испаряемость по НОАК']).style(style);
				worksheet.cell(i + 2, 21).string(objects[i].table2['Индекс вязкости']).style(style);
				worksheet.cell(i + 2, 22).string(objects[i].table2['Вязкость кинематическая при 40°С']).style(style);
				worksheet.cell(i + 2, 23).string(objects[i].table2['Вязкость кинематическая при 100°С']).style(style);
				worksheet.cell(i + 2, 24).string(objects[i].table2['Вязкость кажущаяся (динамическая), определяемая на имитаторе холодной прокрутки (CCS) при -30С']).style(style);
				worksheet.cell(i + 2, 25).string(objects[i].table2['Плотность при +15 цельсия']).style(style);
				worksheet.cell(i + 2, 26).string(objects[i].table2['Температура застывания']).style(style);
				worksheet.cell(i + 2, 27).string(objects[i].table2['Температура вспышки']).style(style);
				worksheet.cell(i + 2, 28).string(objects[i].table2['Температура потери текучести']).style(style);
				worksheet.cell(i + 2, 29).string(objects[i].table2['Общее щелочное число (TBN)']).style(style);
				worksheet.cell(i + 2, 30).string(objects[i].table2['Общее кислотное число (TAN)']).style(style);
				worksheet.cell(i + 2, 31).string(objects[i].table2['Массовая доля серы']).style(style);
				worksheet.cell(i + 2, 32).string(objects[i].table2['Зола сульфатная']).style(style);
				worksheet.cell(i + 2, 33).string(objects[i].table2['Содержание цинка']).style(style);
				worksheet.cell(i + 2, 34).string(objects[i].table2['Содержание фосфора']).style(style);
				worksheet.cell(i + 2, 35).string(objects[i].table2['Содержание бора']).style(style);
				worksheet.cell(i + 2, 36).string(objects[i].table2['Содержание магния']).style(style);
				worksheet.cell(i + 2, 37).string(objects[i].table2['Содержание кальция']).style(style);
				worksheet.cell(i + 2, 38).string(objects[i].table2['Содержание кремния']).style(style);
				worksheet.cell(i + 2, 39).string(objects[i].table2['Содержание натрия']).style(style);
			} catch (err) {
				console.log(objects[i]);
				throw err;
			}
		}
		workbook.write(FOLDER + '/Excel.xlsx');
	} catch (err) {
		throw err;
	}
}

async function sendReq(urlArray, parse) {
	try {
		let fullArray = [];
		let promiseArray = [];
		let bar = new Progress(urlArray.length);
		let max = math.floor(urlArray.length / REQ_IN_TIME);
		console.log('Start sending requests');
		for (let i = 0; i < max; i++) {
			for (let j = i * REQ_IN_TIME; j < (i + 1) * REQ_IN_TIME; j++) {
				promiseArray.push(makeRequest(urlArray[j], parse));
			}
			fullArray = fullArray.concat(await Promise.all(promiseArray));
			promiseArray = [];
			bar.tick(REQ_IN_TIME);
		}
		for (let i = max * REQ_IN_TIME; i < urlArray.length; i++) {
			promiseArray.push(makeRequest(urlArray[i], parse));
		}
		fullArray = fullArray.concat(await Promise.all(promiseArray));
		bar.tick();
		return fullArray;
	} catch (err) {
		console.warn(err);
	}
}

function make() {
	let refs = [];
	for (let i = 0; i < 58; i++) {
		refs.push('https://motoroil24.ru/catalog/oils/motornye/?PAGEN_3=' + (i + 1))
	}
	for (let i = 0; i < 37; i++) {
		refs.push('https://motoroil24.ru/catalog/oils/transmissionnye/?PAGEN_3=' + (i + 1))
	}
	return refs;
}

function normalize(obj) {
	let table1 = {
		'Страна производства': '',
		'Производители масел': '',
		'По производителю автомобиля': '',
		'Область применения': '',
		'Тип продукта': '',
		'Вязкость по SAE': '',
		'API': '',
		'ACEA': '',
		'Спецификации производителей автомобилей': '',
		'Допуски производителей автомобилей': '',
		'Тип двигателя': '',
		'Тип топлива': '',
		'Срок годности, мес': '',
	};
	let table2 = {
		'Внешний вид масла, смазки': '',
		'Цвет продукта': '',
		'Испаряемость по НОАК': '',
		'Индекс вязкости': '',
		'Вязкость кинематическая при 40°С': '',
		'Вязкость кинематическая при 100°С': '',
		'Вязкость кажущаяся (динамическая), определяемая на имитаторе холодной прокрутки (CCS) при -30С': '',
		'Плотность при +15 цельсия': '',
		'Температура застывания': '',
		'Температура вспышки': '',
		'Температура потери текучести': '',
		'Общее щелочное число (TBN)': '',
		'Общее кислотное число (TAN)': '',
		'Массовая доля серы': '',
		'Зола сульфатная': '',
		'Содержание цинка': '',
		'Содержание фосфора': '',
		'Содержание бора': '',
		'Содержание магния': '',
		'Содержание кальция': '',
		'Содержание кремния': '',
		'Содержание натрия': ''
	};
	let result = {
		"naimen": '', "articul": '', "volume": '', "descr": '', "specialist": '', "table1": table1, "table2": table2
	};
	for (let key1 in obj) {
		if (obj.hasOwnProperty(key1) && result.hasOwnProperty(key1) && typeof result[key1] !== 'object' && typeof obj[key1] !== 'undefined') {
			result[key1] = obj[key1];
		}
	}
	for (let key2 in obj.table1) {
		if (obj.table1.hasOwnProperty(key2) && table1.hasOwnProperty(key2) && typeof table1[key2] !== 'object' && typeof obj.table1[key2] !== 'undefined') {
			table1[key2] = obj.table1[key2];
		}
	}
	for (let key3 in obj.table2) {
		if (obj.table2.hasOwnProperty(key3) && table2.hasOwnProperty(key3) && typeof table2[key3] !== 'object' && typeof obj.table2[key3] !== 'undefined') {
			table2[key3] = obj.table2[key3];
		}
	}
	return result;
}

function makeRequestHttp(url, name) {
	return new Promise((resolve, reject) => {
		let file = fs.createWriteStream(name);
		https.get(url, function (response, err) {
			if (response.statusCode === 200) {
				response.pipe(file);
			} else {
				makeRequestHttp(url, name);
			}
			if (err) {
				makeRequestHttp(url, name);
			}
		});
		file.on('finish', () => {
			resolve();
		});
		file.on('error', () => {
			makeRequestHttp(url, name);
		})
	});
}

async function sendReqHttp(urlArray) {
	try {
		let fullArray = [];
		let promiseArray = [];
		let max = math.floor(urlArray.length / REQ_IN_TIME);
		console.log('Start sending requests');
		for (let i = 0; i < max; i++) {
			for (let j = i * REQ_IN_TIME; j < (i + 1) * REQ_IN_TIME; j++) {
				promiseArray.push(makeRequestHttp(urlArray[j].url, urlArray[j].name));
			}
			console.log(math.floor(((i + 1) * REQ_IN_TIME / urlArray.length) * 100));
			fullArray = fullArray.concat(await Promise.all(promiseArray));
			promiseArray = [];
		}
		for (let i = max * REQ_IN_TIME; i < urlArray.length; i++) {
			promiseArray.push(makeRequestHttp(urlArray[i].url, urlArray[i].name));
		}
		fullArray = fullArray.concat(await Promise.all(promiseArray));
		console.log('Success!');
		return fullArray;
	} catch (err) {
		console.warn(err);
	}
}

async function main() {
	try {
		if (!fs.existsSync('./' + FOLDER)) {
			fs.mkdirSync(FOLDER);
		}
		if (!fs.existsSync('./' + FOLDER + '/img')) {
			fs.mkdirSync(FOLDER + '/img');
		}

		let siteHeader = 'https://motoroil24.ru';
		/*		let startRefs = fs.readFileSync('pagesRefs.txt', 'utf-8').split('\r\n');
				let refs = await sendReq(startRefs, function($){
					let references = [];
					let elems = $('body > div.main > div.spanning > div.wrap > div > div.product-item-list')
					.find($('.product.product-item.tovar.js-product'));
					elems.each((i, item)=>{
						let href = $(item).find('a').attr('href').replace(/[/0-9]+$/g, '');
						let offers = $(item).find($('.pack-list')).clone().children();
						offers.each((j, offer)=>{
							references.push(siteHeader + href + $(offer).attr('data-offerid'))
						});
					});
					return references;
				});
				let fullArray = [];
				refs.forEach((item)=> {
					fullArray = fullArray.concat(item);
				});*/
		let startRefs = fs.readFileSync('objectsRefs.txt', 'utf-8').split('\r\n');
		let refs = await sendReq(startRefs, function ($) {
			let result = {};
			result.naimen = $('body > div.main > div.spanning > div.breadcrumbs-box.bb-image > div > div > div.page-title >' + ' h1 > span').text();
			result.articul = $('body > div.main > div.spanning > div.wrap > div > div.tovar > div.product-description >' + ' div.l-status__item.l-code__item-detail').text().replace('Код продукта: ', '');
			result.volume = $('body > div.main > div.spanning > div.wrap > div > div.tovar > div.product-description >' + ' div.product-param_block.clearfix > div.product-param_item.pack > div.product-param_item-value').text();
			result.descr = $('body > div.main > div.spanning > div.wrap > div > div.tovar > div:nth-child(7) >' + ' div').text();
			result.specialist = $('body > div.main > div.spanning > div.wrap > div > div.tovar > div:nth-child(8) >' + ' div').text();
			let table1 = $('body > div.main > div.spanning > div.wrap > div > div.tovar')
			.find('table.b-harak_table:nth-child(3)').clone().children();
			let table2 = $('body > div.main > div.spanning > div.wrap > div > div.tovar')
			.find('table.b-harak_table:nth-child(5)').clone().children();
			let resultTable1 = [];
			let resultTable2 = [];
			table1.each((i, item) => {
				$(item).find('tr').each((j, itemJ) => {
					let head = $(itemJ)
					.find('td:nth-child(1)')
					.text()
					.replace(/[\n\r]+/g, '')
					.replace(/[ \t]+/g, ' ')
					.replace(/ $/g, '');
					resultTable1[head] = $(itemJ)
					.find('td:nth-child(2)')
					.text()
					.replace(/[\n\r]+/g, '')
					.replace(/[ \t]+/g, ' ');
				});
			});
			table2.each((i, item) => {
				$(item).find('tr').each((j, itemJ) => {
					let head = $(itemJ)
					.find('td:nth-child(1)')
					.text()
					.replace(/[\n\r]+/g, '')
					.replace(/[ \t]+/g, ' ')
					.replace(/ $/g, '');
					resultTable2[head] = $(itemJ)
					.find('td:nth-child(2)')
					.text()
					.replace(/[\n\r]+/g, '')
					.replace(/[ \t]+/g, ' ');
				});
			});
			result.table1 = resultTable1;
			result.table2 = resultTable2;
			let imgHref = $('body > div.main > div.spanning > div.wrap > div > div.tovar >' + ' div.product-gallery-box > div.product-gallery').find('img').attr('src');
			let docFabrHref = $('body > div.main > div.spanning > div.wrap > div > div.tovar >' + ' div.product-description > div.product-docs_block > a.b-upload__item').attr('href');
			let secPassHref = $('body > div.main > div.spanning > div.wrap > div > div.tovar >' + ' div.product-description > div.product-docs_block > a.b-upload__item2').attr('href');
			if (typeof imgHref !== 'undefined') {
				result.imgHref = siteHeader + imgHref;
				result.imgName = result.articul + '.jpeg';
			}
			if (typeof docFabrHref !== 'undefined') {
				result.docFabrHref = siteHeader + docFabrHref;
				result.docFabrName = 'Паспорт производителя ' + result.articul + '.pdf';
			}
			if (typeof secPassHref !== 'undefined') {
				result.secPassHref = siteHeader + secPassHref;
				result.secPassName = 'Паспорт безопастности ' + result.articul + '.pdf';
			}
			return result;
		});
		let files = [];
		refs.forEach((item) => {
			if (typeof files[item.table1['Производители масел']] === 'undefined') {
				files[item.table1['Производители масел']] = {};
				if (!fs.existsSync('./' + FOLDER + '/img/' + item.table1['Производители масел'])) {
					fs.mkdirSync('./' + FOLDER + '/img/' + item.table1['Производители масел']);
				}
			}
			if (typeof item.articul !== 'undefined') {
				files[item.table1['Производители масел']][item.articul] = [];
				if (typeof item.imgHref !== 'undefined') {
					files.push({
						name: './' + FOLDER + '/img/' + item.table1['Производители масел'] + '/' + item.imgName,
						url: item.imgHref
					})
				}
				if (typeof item.docFabrHref !== 'undefined') {
					files.push({
						name: './' + FOLDER + '/img/' + item.table1['Производители масел'] + '/' + item.docFabrName,
						url: item.docFabrHref
					})
				}
				if (typeof item.secPassHref !== 'undefined') {
					files.push({
						name: './' + FOLDER + '/img/' + item.table1['Производители масел'] + '/' + item.secPassName,
						url: item.secPassHref
					})
				}
			}
		});
		fs.writeFileSync('1.json', JSON.stringify(files, undefined, 4));
		refs = refs.map((item) => {
			return normalize(item);
		});
		console.log(refs.length);
		saveAsXLSX(refs);
		doit();
	} catch (err) {
		console.log('main', err)
	}
}

function download(url, path) {
	return new Promise((resolve, reject) => {
		let options = {
			method: 'GET', uri: url, resolveWithFullResponse: true
		};
		let file = fs.createWriteStream(path);
		rp(options).pipe(file);
		file.on('finish', () => {
			resolve();
		});
		file.on('error', () => {
			console.log('file error;');
			download(url, path);
		});
	});
}

async function doit() {
	try {
		let data = fs.readFileSync('1.json');
		let json = JSON.parse(data);
		for (let i = 0; i < json.length; i++) {
			await download(json[i].url, json[i].name);
			console.log(json.length, i, json[i].name);
		}
		console.log('end;');
	} catch (err) {
		console.log('ERR');
	}
};

main();
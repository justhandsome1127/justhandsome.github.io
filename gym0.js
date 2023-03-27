/*
var date1,date2;
date1 = new Date("03/15/2023"); 
date2 = new Date();
 
var time_difference = date2.getTime() - date1.getTime();  
var days = time_difference / (1000 * 60 * 60 * 24); 
days = Math.ceil(days);

var price = 30*days;

let minutes = window.prompt("輸入分鐘");

minutes = Number(minutes);
 
alert("每分鐘" +  price / minutes + "元");
*/


var date1,date2;
date1 = new Date("03/15/2023"); 
date2 = new Date();
 
var time_difference = date2.getTime() - date1.getTime();  
var days = time_difference / (1000 * 60 * 60 * 24); 
days = Math.ceil(days);

var price = 30*days; 
var minutes = 0;

var XLSX = require("xlsx");

var workbook = XLSX.readFile("gym.xlsx");

let worksheet = workbook.Sheets["Sheet1"];

for(let index = 2; index <= days+1; index++){
    let temp;
    temp = worksheet[`G${index}`].v;
    minutes += Number(temp);
}

console.log("你平均一分鐘花：" + price/minutes + "元");



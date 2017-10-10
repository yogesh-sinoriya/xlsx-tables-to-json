var XLSX = require('xlsx');


module.exports = function(path){

	var exceltojson, converted_data;
	var workbook = XLSX.readFile(path);
	var sheetnames_list = workbook.SheetNames;
	var expacted_file = false;
	var data = {};
	var headers = {};
	var tables = "";
	var dyn_row = 0;
	sheetnames_list.forEach(function(y) {
		var worksheet = workbook.Sheets[y];
		data[y] = {};
		try{
			expacted_file = true;


		for(ar in worksheet) {
			if(ar[0] === '!') continue;
			var tt = 0;
			for (var i = 0; i < ar.length; i++) {
				if (!isNaN(ar[i])) {
					tt = i;
					break;
				}
			};
			var col = ar.substring(0,tt);
			var row = parseInt(ar.substring(tt));
			var value = worksheet[ar].v;

			//console.log(col+row+'--->'+value);
			worksheet['!merges'].forEach(function(r) {
					if(row === r.s.r+1){
					//console.log(row);

					tables = value;
					dyn_row = row+1;
					data[y][value] = [];
					headers = {};
					tempRow = 0;

					}
			});

			if(row == dyn_row && value) {
						headers[col] = value;
						continue;
			}

			if(row > dyn_row){

				if(!data[y][tables][row]) data[y][tables][row]={};
				data[y][tables][row][headers[col]] = value;

				if(!headers[col]){
					delete data[y][tables][row];
				}

			}
		}
		}catch(err){
			//console.log(err);
				expacted_file = false;
		}
	});

	data = remove_empty(data);
	//console.log(data)
	//console.log(JSON.stringify(data));
	if(expacted_file){
		return data;
	}else{
		//console.error(new Error('Format not supported!'));
		return {error:'Format not supported'};

	}

};

var remove_empty = function ( data ) {

	var SheetNames = Object.keys(data);
	SheetNames.forEach(function(SheetName){
		var TableNames = Object.keys(data[SheetName]);
		TableNames.forEach(function(TableName){
			//console.log(`---------------------DB ${SheetName} Table ${TablqeName}------------------------`);
			data[SheetName][TableName] = trim_nulls(data[SheetName][TableName]);
			//console.log(data[SheetName][TableName]);

		});
	});

	return data;

};

var trim_nulls = function (data) {

	var keys = Object.keys(data);
	var y;
	//console.log(keys);
	keys.forEach(function(key,index){
		data[index] = data[key];
		delete data[key];
	});

	for(var x=0;x<data.length;x++){
		y = data[x];
		if (y==="null" || y===null || y==="" || typeof y === "undefined" || (y instanceof Object && Object.keys(y).length == 0)) {
			//delete data[x];
			data.splice(x , data.length);
		}
	}
	//data.clean(undefined);
	return data;
}

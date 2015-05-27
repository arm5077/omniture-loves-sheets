var express = require("express");
var app = express();
var cron = require('cron').CronJob;
var report = require('nomniture').Report;
var Spreadsheet = require('edit-google-spreadsheet');

var spreadsheet_email = process.env.SPREADSHEET_EMAIL;
var spreadsheet_key = process.env.SPREADSHEET_KEY;

// Thanks to http://stackoverflow.com/questions/3066586/get-string-in-yyyymmdd-format-from-js-date-object for this code
Date.prototype.yyyymmdd = function() {
   	var yyyy = this.getFullYear().toString();
   	var mm = (this.getMonth()+1).toString(); // getMonth() is zero-based
   	var dd  = this.getDate().toString();
	return yyyy + "-" + (mm[1]?mm:"0"+mm[0]) + "-" + (dd[1]?dd:"0"+dd[0]); // padding
};


// Initialize Omniture login
var r = new report(process.env.USERNAME, process.env.SECRET, process.env.SERVER, {
	waitTime:5,
	log:false
});


// Using this until I can figure out what's wrong with cron
setInterval(update, 60*60*1000);

app.get("/", function(request, response){
	update();
	response.status(200).json("Updating spreadsheet now!")
});

var port = process.env.PORT || 3000;
app.listen(port, function(){
	console.log("We're live at port " + port + ".");
});


function update(){
		Spreadsheet.load({
			debug: true,
			spreadsheetName: "Graphics projects",
			worksheetName: "Projects",
			oauth: {
				email: spreadsheet_email,
				key: spreadsheet_key
			}

		}, function(err, spreadsheet){
			if( err ) throw err;
			spreadsheet.receive(function(err, rows, info) {
				if( err ) throw err;

				// Cycle through spreadsheet and create new object
				var graphics = [];
				var fields = [];

				for( var row in rows ){
					if( rows.hasOwnProperty(row) ){
						var object = {};
						for( var column in rows[row] ){
							if( rows[row].hasOwnProperty(column) ){
								if( row == "1" ){
									fields.push(rows[row][column]);
								}
								else {
									object[fields[column - 1]] = rows[row][column];
								}
							}
						}
						graphics.push(object);
					}
				}

				var metrics = [];
				// Search through field list and find requested Omniture metrics
				fields.forEach(function(field){
					if( field.search("GET: ") != -1 ) {
						metrics.push( { "id": field.substring( "GET: ".length ) });
					}
				});

				// OK, now let's cycle through this bad boy and get some traffic numbers
				graphics.forEach(function(graphic, index){
					if( graphic["Finished headline"] ){

						console.log("Pulling stats for " + graphic["Finished headline"] + "...")

						var yearAgo = new Date();
		 				yearAgo.setYear( yearAgo.getFullYear() - 1 );
						
						r.request("Report.QueueRanked", {
							"reportDescription": {
						    "reportSuiteID": process.env.REPORTSUITEID,
						    "dateFrom": yearAgo.yyyymmdd(),
						    "dateTo": new Date().yyyymmdd(),
						    "metrics": metrics,
						    "elements": [
						      {
						        "id": "page",
						        "top": 5,
						        "search": {
						          "type": "OR",
						          "keywords": [graphic["Finished headline"].replace(/'/g, '*')]
						        }
						      }
						    ]
						  }
						}, function(err, response){
							if(err) throw err;

							console.log("Found stats for " + graphic["Finished headline"] + ": " + response)

							response = JSON.parse(response);
							metrics.forEach(function(metric, subIndex){
								var exportColumn = {}, exportRow = {};
								response.report.data.forEach(function(datum){
									var column = (fields.indexOf("GET: " + metric.id) + 1).toString();
									var row = index + 1;
									exportColumn[column] = (!exportColumn.hasOwnProperty(column) ? parseInt(datum.counts[subIndex]) : exportColumn[column] + parseInt(datum.counts[subIndex]));
									console.log("THIS IS THE NUMBER WE'RE DEALING WITH: " + parseInt(datum.counts[subIndex]));
									exportRow[row] = exportColumn;
								});
								console.log("Sending this to spreadsheet (re: " + graphic["Finished headline"] + "): " + JSON.stringify(exportRow));
								spreadsheet.add(exportRow);
							});
							
								
						
							spreadsheet.send(function(err) {
								if(err) throw err;
							});	
						});	
					}		
				});
		    });
		});
}


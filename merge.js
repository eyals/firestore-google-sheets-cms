import fs from 'fs';

const srcPath = 'src';
const outputFile = 'dist/firestore-sheets-cms.gs';
const excludeFiles = ['appsscript.json'];
const scriptExtension = 'js';


fs.writeFile(outputFile, '', ()=>console.log('Destination cleared'))


fs.readdir(srcPath, function (err, files) {
	if (err) {
		return console.log('Unable to scan directory: ' + err);
	}
	// Excluding files from excludeFiles or without scriptExtension
	let scriptDocs = files.filter(f =>
		excludeFiles.indexOf(f) == -1 &&
		f.substring(f.lastIndexOf(".") + 1) == scriptExtension
	);

	scriptDocs.forEach((doc) => {
		readAppend(outputFile, [srcPath,doc].join("/"));
	});
	console.log("after all");
});

function readAppend(destination, file) {
	fs.readFile(file, 'utf8', (err, data) => {
		if (err) throw err;
		console.log("File read", file);
		fs.appendFile(destination, data, (err) => {
			if (err) throw err;
			console.log("Destination updated");
		});
	});

}

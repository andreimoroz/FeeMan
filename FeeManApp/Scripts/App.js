function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}
 
function s2ab(s) {
	var buf = new ArrayBuffer(s.length);
	var view = new Uint8Array(buf);
	for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

function getLookup(vals, name) {
	var val = vals[name];
	if (val) { val = val.get_lookupValue(); }
	return val;
}

function saveExcel(s) {
	// original data
//	var data = [[1,2,3],[true, false, null, "sheetjs"],["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"], ["baz", null, "qux"]]

	try {
		var ctx = SP.ClientContext.get_current();
		
		// load items
		var camlQuery = new SP.CamlQuery();
		camlQuery.set_viewXml(
			'<View><Query>' +
					'<Eq>' +
						'<FieldRef Name="fmWorkflowActive"></FieldRef>' +
						'<Value Type="Text">-</Value>' +
					'</Eq>' +
					'<OrderBy> <FieldRef Name="ID"  Ascending="TRUE/></OrderBy>' +
					'<ViewFields><FieldRef Name="ID"/><FieldRef Name="Author"/><FieldRef Name="fmTeam"/><FieldRef Name="fmAccount"/><FieldRef Name="fmClientName"/><FieldRef Name="fmFeePLVDescription"/><FieldRef Name="fmApprovedAt"/><FieldRef Name="fmUnits"/></ViewFields>' +
			'</Query></View>'
		);

		var listFees = ctx.get_web().get_lists().getByTitle('Fees');
		var listItems = listFees.getItems(camlQuery);
		ctx.load(listItems);
		
		ctx.executeQueryAsync(
			function () {
				if (listItems.get_count() == 0)
				{
					alert("Nothing to export");
					return;
				}
				 
				var ws = {};
				var range = {s: {c:0, r:0}, e: {c:19, r: listItems.get_count() }};
				ws['A1'] = {v: "User", t: "s" }; // Author
				ws['B1'] = {v: "User's team", t: 's' }; // fmTeam
				ws['C1'] = {v: "Service Unique Reference Code", t: 's' }; // ID
				ws['D1'] = {v: "Source", t: 's' }; // 'SP'
				ws['E1'] = {v: "Contract/Structure", t: 's' };
				ws['F1'] = {v: "Client Name", t: 's' }; // fmClientName
				ws['G1'] = {v: "Branch Code", t: 's' }; // fmAccount 0-4
				ws['H1'] = {v: "Radix Number of the Client", t: 's' }; // fmAccount 6-11
				ws['I1'] = {v: "Ordinal of account", t: 's' }; // fmAccount 13-15
				ws['J1'] = {v: "Key of account", t: 's' }; // fmAccount 17-18
				ws['K1'] = {v: "Account currency", t: 's' }; // fmAccount 20-22
				ws['L1'] = {v: "IBAN Account", t: 's' };
				ws['M1'] = {v: "Description", t: 's' };
				ws['N1'] = {v: "Charge code", t: 's' }; // feePlvDescription ?
				ws['O1'] = {v: "Event date", t: 's' }; // fmApprovedAt
				ws['P1'] = {v: "Number of Units", t: 's' }; // fmUnits
				ws['Q1'] = {v: "Transaction amount", t: 's' };
				ws['R1'] = {v: "Transaction currency", t: 's' };
				ws['S1'] = {v: "Fee amount", t: 's' };
				ws['T1'] = {v: "Fee currency", t: 's' };

				var row = 2;
				var enm = listItems.getEnumerator();
                while (enm.moveNext()) {
                    var item = enm.get_current();
					var vals = item.get_fieldValues();
					var acc = vals["fmAccount"];
					ws['A' + row] = {v: getLookup(vals, "Author"), t: 's' }
					ws['B' + row] = {v: getLookup(vals, "fmTeam"), t: 's' }
					ws['C' + row] = {v: "" + vals["ID"], t: 's' }
					ws['D' + row] = {v: "SP", t: 's' };
					ws['F' + row] = {v: vals["fmClientName"], t: 's' }
					ws['G' + row] = {v: acc.substring(0, 5), t: 's' }; // fmAccount 0-4
					ws['H' + row] = {v: acc.substring(6, 12), t: 's' }; // fmAccount 6-11
					ws['I' + row] = {v: acc.substring(13, 16), t: 's' }; // fmAccount 13-15
					ws['J' + row] = {v: acc.substring(17, 19), t: 's' }; // fmAccount 17-18
					ws['K' + row] = {v: acc.substring(20, 23), t: 's' }; // fmAccount 20-22
					ws['L' + row] = {v: "", t: 's' };
					ws['M' + row] = {v: "", t: 's' };
					ws['N' + row] = {v: getLookup(vals, "fmFeePLVDescription"), t: 's' };
					ws['O' + row] = {v: vals["fmApprovedAt"].format("dd-MM-yyyy"), t: 's' };
					ws['P' + row] = {v: "" + vals["fmUnits"], t: 's' };
					ws['Q' + row] = {v: "", t: 's' };
					ws['R' + row] = {v: "", t: 's' };
					ws['S' + row] = {v: "", t: 's' };
					ws['T' + row] = {v: "", t: 's' };
//					alert("Acc: " + acc);
					row++;
                }

				ws['!ref'] = XLSX.utils.encode_range(range);
				 
				// add worksheet to workbook
				var wb = new Workbook();
				var ws_name = "SheetJS";
				wb.SheetNames.push(ws_name);
				wb.Sheets[ws_name] = ws;
				var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
				
				saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "FeeMan." + (new Date).format('yyyy-MM-dd.HH-mm') + ".xlsx");
			},
            function (message) {
                alert("ERROR: " + message);
            });
	} catch (e) {
		console.log("Error: " + e);
		alert("ERROR: " + e.message);
//		CIB.logging.logError('error', e.message);
	}
}

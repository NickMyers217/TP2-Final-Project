/*
 * This is Nick's Form Validation Javascript
 * (DO NOT COPY THIS, JERK)
 */
function Journal(form) {
	this.form = form;
	this.srn = {
		value: $(form + " #srn").val(),
		regEx: /^[0-9]{4}$/
	};
	this.month = {
		value: this.pad($(form + " #month").val(), 2, "0"),
		regEx: /^(0[1-9]|1[0-2])$/
	};
	this.gl = {
		data: [],
		acctRegEx: {
			zeroes: /^[0]{16}$/,
			four: /^[0-9]{4}$/,
			currency: /^[-+]?[0-9]*\.?[0-9]{1,2}$/
		}
	};
	this.xml = "";
} Journal.prototype = {
	constructor: Journal,
	formToArray: function() {
		for(var i = 1; i <= 36; i += 6) {
			var glLine = [];
			if(!this.checkRow(i, 6))
				continue;
			for(var j = 0; j < 6; j++) {
				var thisVal = $(this.form + (i + j)).val();
				if(j < 4 && thisVal.length > 0 && thisVal.length < 4)
					thisVal = this.pad(thisVal, 4, "0");	
				glLine.push(thisVal);
			}
			this.gl.data.push(glLine);
		}	
	},
	checkRow: function(row, cols) {
		for(var col = 0; col < cols; col++)
			if($(this.form + (row + col)).val() != "")
				return true;
		return false;
	},
	pad: function(value, target, symbol) {
		var thisVal = value;
		while(thisVal.length < target)
			thisVal = symbol + thisVal
		return thisVal;
	},
	calculateSum: function() {
		var sum = 0.00;
		for(var i = 0; i < this.gl.data.length; i++)
			sum += parseFloat(this.gl.data[i][4]);
		return sum.toFixed(2);
	},
	arrayToXML: function() {
		var xml = "<journalvoucher><srn>" + this.srn.value + "</srn>";
		xml += "<fm>" + this.month.value + "</fm>";
		for(var i = 0; i < this.gl.data.length; i++) {
			xml += "<je>";
			xml += "<major>" + this.gl.data[i][0] + "</major>";
			xml += "<minor>" + this.gl.data[i][1] + "</minor>";
			xml += "<sub1>" + this.gl.data[i][2] + "</sub1>";
			xml += "<sub2>" + this.gl.data[i][3] + "</sub2>";
			xml += "<ta>" + this.gl.data[i][4] + "</ta>";
			xml += "<desc>" + this.gl.data[i][5] + "</desc>";
			xml += "</je>";
		}
		xml += "</journalvoucher>";
		this.xml = xml;
	},
	submitXML: function() {
		$("#xmlout").val(this.xml);
		document.forms[2].submit();
	}
};

function Errors() {
	this.srn = [];
	this.month = [];
	this.gl = {
		acctID: [],
		ta: []
	};
	this.sum = [];
} Errors.prototype = {
	constructor: Errors,
	count: function() {
		var c = 0;
		c += this.srn.length;
		c += this.month.length;
		c += this.gl.acctID.length;
		c += this.gl.ta.length;
		c += this.sum.length;
		return c;
	},
	toHTML: function() {
		var os = "<p>Errors:<br><ol>";
		for(var key in this) {
			if(key == "srn" || key == "month" || key == "sum") {
				var arr = this[key];
				if(arr.length > 0)
					for(var i = 0; i < arr.length; i++)
						os += "<li>" + arr[i] + "</li>";
			}
			if(key == "gl") {
				for(var key in this.gl) {
					var arr = this.gl[key];
					if(arr.length > 0)
						for(var i = 0; i < arr.length; i++)
							os += "<li>" + arr[i] + "</li>";
				}
			}
		}
		os += "</ol>";
		return os;
	}
};

function Validate(form, errors) {
	this.form = form;
	this.errors = errors;
} Validate.prototype = {
	constructor: Validate,
	validSRN: function() {
		if(!(this.form.srn.regEx.test(this.form.srn.value)))
			this.errors.srn.push("The SRN must be a valid 4 digit number.");
	},
	validMonth: function() {
		if(!(this.form.month.regEx.test(this.form.month.value)))
			this.errors.month.push("The fiscal month must be a number 1 through 12.");
	},
	validGl: function() {
		var gl = this.form.gl;
		var data = gl.data;
		if(data.length < 2) {
			this.errors.gl.acctID.push("2 or more general ledger entries must be made.");
			return;
		}
		for(var i = 0; i < data.length; i++) {
			var acctID = data[i][0] + data[i][1] + data[i][2] + data[i][3]
			if(gl.acctRegEx.zeroes.test(acctID)) {
				this.errors.gl.acctID.push("General ledger entry " + (i+1) + " cannot be zero.");
				continue;
			}
			for(var j = 0; j < data[i].length; j++) {
				if(j < 4) {
					if(!(gl.acctRegEx.four.test(data[i][j]))) {
						this.errors.gl.acctID.push("General ledger entry " + (i+1) + " contains a field that is not a valid 4 digit number.");
						break;
					}
				}
			}
			if(this.blank(data[i][4])) {
				this.errors.gl.ta.push("General ledger entry " + (i+1) + " has no transaction amount.");
				continue;
			} else {
				if(!(gl.acctRegEx.currency.test(data[i][4])))
					this.errors.gl.ta.push(data[i][4] + " is not a valid currency amount.");
			}
		}
	},
	validSum: function() {
		if(this.errors.gl.ta.length == 0) {
			var sum = this.form.calculateSum();
			if(sum != 0.00)
				this.errors.sum.push("The sum did not total to zero.");
		}
	},
	blank: function(input) {
		if(input == "") return true;
		else return false;
	}
};

function main() {
	var journal = new Journal("#gl");
	var errors = new Errors();
	var validate = new Validate(journal, errors);

	journal.formToArray();
	validate.validSRN();
	validate.validMonth();
	validate.validGl();
	validate.validSum();

	console.log(journal);
	console.log(errors);

	if(errors.count() == 0) {
		journal.arrayToXML();
		journal.submitXML();
	} else {
		errorHTML = errors.toHTML();
		$("#errors").html(errorHTML);
	}
}

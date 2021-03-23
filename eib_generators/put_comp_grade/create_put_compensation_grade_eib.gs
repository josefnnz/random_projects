// create alphabet for column references
var NUM_COLUMNS = 300;
var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
var column_letters = alphabet;
for (var i = 26; i < NUM_COLUMNS; i++) {
	var quotient = Math.trunc(i / 26);
	var remainder = i % 26;
	column_letters[i] = column_letters[quotient - 1] + column_letters[remainder];
}

var column_letter_indices = {};
for (var i = 0; i < column_letters.length; i++) {
	column_letter_indices[column_letters[i]] = i;
}
/**
 * This function replaces template tags with data
 *
 * @param {String} template text including the placeholder tags
 * @param {String} data is the actual data to replace the placeholders in the template
 * @return {String} template is returned after modification
 */
function replaceVariables_(template, data) {
	// {{email address}}
	var templateVars = template.match(/\{\{(?:[^\}\}]+)\}\}/g);
	if (templateVars != null) {
		for (var i = 0; i < templateVars.length; ++i) {
			var text = data[normalizeHeader_(templateVars[i])] || "";
			if (text instanceof Date) {
				var timestamp = Date.parse(text)
				if (isNaN(timestamp) == false) {
					text = Utilities
							.formatDate(new Date(timestamp), SpreadsheetApp
									.getActive().getSpreadsheetTimeZone(),
									"MMM d, YYYY");
				}
			}
			template = template.replace(templateVars[i], text);
		}
	}
	return template;
}

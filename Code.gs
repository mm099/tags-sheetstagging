function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var activeCell = ss.getActiveCell();

  var newValue = String(e.value);
  var oldValue = String(e.oldValue);

  var tagColumnName = 'Tags'
  var tagColumnSheet = 1

  if (activeSheet.getRange(1, activeCell.getColumn()).getValue() === tagColumnName && activeSheet.getIndex() === tagColumnSheet) {
    var rule = activeCell.getDataValidation();
    var args = rule.getCriteriaValues()[0];

    // new value is empty
    if (!newValue || newValue === '' || newValue === 'Empty') {
      newValue = '';
    }
    else {
      var newValueSplit = newValue.split('; ');
      // the user manually replaces the current value with a list of tags
      if (newValueSplit.length > 1) {
        // validate each new tag and sort
        var newValueValidated = newValueSplit.filter(function(element) {
          return args.includes(element);
        });
        newValue = newValueValidated.sort().join('; ');
      }
      else {
        // first of all, remove invalid values from current value
        oldValue = oldValue.split('; ').filter(function(element) {
          return args.includes(element);
        }).join('; ');

        // single new tag is selected for blank cell
        if (!e.oldValue || oldValue === '') {
          // make changes only if the new tag is valid
          if (!args.includes(newValue)) {
            newValue = '';
          }
          else {
            newValue = newValue;
          }
        }

        // single new tag selected for non-blank cell
        else {

          // make changes only if the new tag is valid
          if (!args.includes(newValue)) {
            newValue = oldValue;
          }

          // add the tag only if it is not a duplicate of an existing tag
          else if(oldValue.indexOf(newValue) < 0) {          
            // insert new tag into list alphabetically
            let newValueArray = oldValue.split('; ');
            let i = newValueArray.findIndex(function(element) {
              return element.localeCompare(newValue) === 1;
            });
            if (i === -1) {
              i = newValueArray.length;
            }
            newValueArray.splice(i, 0, newValue)
            newValue = newValueArray.join('; ');
          }
          
          // otherwise, remove the tag from the list
          else {
            const newValueArray = oldValue.split('; ').filter(function(element) {
                return element !== newValue;
            });
            newValue = newValueArray.join('; ');
          }
        }
      }
    }
    activeCell.setValue(newValue);
  }
}

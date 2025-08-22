/**
 * @fileoverview Updated Form.js for frequency-based habit tracking
 * Manages the Google Form for daily habit completion counting
 */

/**
 * Finds the Google Form linked to the spreadsheet.
 * @return {Form} The Google Form object.
 */
function getForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // First, try to find the form via the correct sheet name.
  let formUrl = null;
  const trackingSheet = ss.getSheetByName(SHEET_NAMES.TRACKING);
  if (trackingSheet) {
    formUrl = trackingSheet.getFormUrl();
  }
  
  // If not found, try the default name as a fallback.
  if (!formUrl) {
    const defaultFormSheet = ss.getSheetByName('Form Responses 1');
    if (defaultFormSheet) {
      formUrl = defaultFormSheet.getFormUrl();
    }
  }

  if (!formUrl) {
    throw new Error('No Google Form linked to this spreadsheet was found. Please ensure a form is linked to the "Daily_Tracking" or "Form Responses 1" sheet.');
  }

  return FormApp.openByUrl(formUrl);
}

/**
 * Updates the 'Habit Selection' dropdown in the Google Form
 * with the list of 'Active' habits from the Habits_Main sheet.
 */
function updateHabitFormDropdown() {
  log('INFO', 'Updating habit form dropdown...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const habitsSheet = ss.getSheetByName(SHEET_NAMES.HABITS);
  
  if (!habitsSheet) {
    log('ERROR', 'Habits_Main sheet not found. Cannot update form dropdown.');
    return;
  }
  
  // Check if there is any data beyond the header row
  if (habitsSheet.getLastRow() < 2) {
    log('WARN', 'Habits_Main sheet is empty. No habits to add to the form dropdown.');
    return;
  }
  
  const habitsData = habitsSheet.getRange(2, 1, habitsSheet.getLastRow() - 1, 2).getValues();
  const activeHabits = habitsData.filter(row => row[0]); // Simple filter for non-empty rows
  
  if (activeHabits.length === 0) {
    log('WARN', 'No active habits found to add to the form dropdown.');
    return;
  }
  
  try {
    const form = getForm();
    const items = form.getItems();
    let habitDropdownItem = null;
    
    items.forEach(item => {
      if (item.getTitle() === 'Habit Selection') {
        habitDropdownItem = item;
      }
    });

    if (!habitDropdownItem) {
      log('ERROR', 'Habit Selection dropdown not found in the form.');
      return;
    }
    
    const habitChoices = activeHabits.map(habit => habit[1]);
    const habitList = habitDropdownItem.asListItem();
    habitList.setChoices(habitChoices.map(choice => habitList.createChoice(choice)));
    
    log('INFO', 'Habit form dropdown updated successfully with', habitChoices.length, 'items.');

  } catch (error) {
    log('ERROR', 'Failed to update form dropdown:', error.message);
  }
}

/**
 * Sets up a form submit trigger and returns the form URL.
 * This is a helper function for initial setup.
 * @return {string} The URL of the Google Form.
 */
function setupFormTrigger() {
  const form = getForm();
  
  // Delete existing triggers to prevent duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create the new trigger
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
  
  log('INFO', 'Form trigger created successfully.');
  return form.getPublishedUrl();
}

/**
 * Creates or updates the Google Form to use frequency-based tracking
 * This should be run after migrating to the frequency approach
 */
function updateFormForFrequencyTracking() {
  try {
    log('INFO', 'Updating form for frequency-based tracking...');
    
    const form = getForm();
    const items = form.getItems();
    
    // Remove or update time-related questions
    items.forEach(item => {
      const title = item.getTitle();
      
      // Remove the time completion question if it exists
      if (title === 'Time of completion' || title.includes('Time')) {
        form.deleteItem(item);
        log('INFO', `Removed time-related question: ${title}`);
      }
    });
    
    // Update the Success/Miss question to be clearer for frequency tracking
    let successItem = null;
    items.forEach(item => {
      if (item.getTitle() === 'Success/Miss' || item.getTitle().includes('Success')) {
        successItem = item;
      }
    });
    
    if (successItem) {
      const choices = ['Completed', 'Missed'];
      const choiceItems = choices.map(choice => successItem.asListItem().createChoice(choice));
      successItem.asListItem().setChoices(choiceItems);
      successItem.setTitle('Completion Status');
      successItem.setHelpText('Did you complete this habit occurrence?');
      log('INFO', 'Updated completion status question');
    }
    
    // Add or update completion count question
    let countItem = null;
    items.forEach(item => {
      if (item.getTitle().includes('Count') || item.getTitle().includes('Times')) {
        countItem = item;
      }
    });
    
    if (!countItem) {
      // Add new completion count question
      countItem = form.addListItem();
      countItem.setTitle('How many times completed today?');
      countItem.setHelpText('Select the number of times you have completed this habit today');
      
      const countChoices = ['1', '2', '3', '4', '5+'];
      const countChoiceItems = countChoices.map(choice => countItem.createChoice(choice));
      countItem.setChoices(countChoiceItems);
      countItem.setRequired(true);
      
      log('INFO', 'Added completion count question');
    }
    
    log('INFO', 'Form updated successfully for frequency tracking');
    
  } catch (error) {
    log('ERROR', 'Failed to update form for frequency tracking:', error.message);
  }
}
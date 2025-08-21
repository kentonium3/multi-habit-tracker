/**
 * @fileoverview Custom function to generate unique habit IDs.
 * This script uses Script Properties to ensure the IDs are always unique,
 * even if rows in the Habits_Main sheet are deleted.
 */

/**
 * Generates a unique, sequential habit ID.
 * This function uses the Script Properties service to maintain a persistent
 * counter, ensuring that each generated ID is unique and follows the
 * "H001", "H002" format.
 * * To use this function, simply type "=generateHabitId()" into the
 * "HabitID" cell in your Habits_Main sheet.
 *
 * @return {string} A unique habit ID (e.g., "H001").
 * @customfunction
 */
function generateHabitId() {
  try {
    // Get the ScriptProperties service, which stores key-value pairs
    // persistently for this specific script project.
    const scriptProperties = PropertiesService.getScriptProperties();

    // Retrieve the last used ID number. The number is stored as a string.
    let lastIdNumber = scriptProperties.getProperty('lastHabitIdNumber');
    
    // If this is the first time the function is run, initialize the counter to 0.
    if (!lastIdNumber) {
      lastIdNumber = 0;
    } else {
      // Convert the string to a number for calculation.
      lastIdNumber = parseInt(lastIdNumber);
    }

    // Increment the counter for the new ID.
    const newIdNumber = lastIdNumber + 1;
    
    // Store the new number back in the script properties for future use.
    scriptProperties.setProperty('lastHabitIdNumber', newIdNumber.toString());

    // Format the number into a three-digit string with a leading 'H'.
    const newId = 'H' + newIdNumber.toString().padStart(3, '0');

    log('DEBUG', `Generated new habit ID: ${newId}`);
    return newId;

  } catch (error) {
    log('ERROR', 'Failed to generate habit ID:', error.message, error.stack);
    // Return an error message to the user in the spreadsheet cell.
    return "ERROR: Failed to generate ID";
  }
}

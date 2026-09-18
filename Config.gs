const SCC_CONFIG = {
  sheets: {
    callEntry: 'Call Entry',
    roster: 'Roster'
  },

  callEntry: {
    studentSelector: 'C2',
    studentId: 'O2',
    note: 'D1',
    toDo: 'I1',
    formRange: 'A7:I50',

    resetRanges: [
      'H14',      // Message left
      'H15',      // New address
      'H17:I22',  // Updated phone/email information
      'C20',      // Learning Coach availability
      'C22',      // Translator language
      'C32:E32',  // Barrier details
      'C36',      // School-year goal
      'C37',      // Interests
      'C39',      // Career/future goal
      'C40',      // Counselor entry
      'H42',      // Course corrections
      'C47:H47'   // Additional information
    ]
  },

  roster: {
    headerRow: 1,

    studentIdHeader: 'Student Number',
    lastNameHeader: 'LAST NAME',
    firstNameHeader: 'FIRST NAME',

    completionHeader: 'SCC Completion',
    notesHeader: 'SCC Notes',
    toDoHeader: 'SCC To Do',

    completedValue: 'Completed'
  }
};

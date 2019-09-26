define([], function() {
  return {
    PropertyPaneGeneralDescription:
      "Configure this web part's general settings",
    PropertyPaneHeaderStylesDescription:
      "Configure header style on this page, select colours from the colour pickers.",
    PropertyPaneQuestionStylesDescription:
      "Configure question style on this page, select colours from the colour pickers.",
    PropertyPaneAnswerStylesDescription:
      "Configure answer style on this page, select colours from the colour pickers.",
    BasicGroupName: "General",
    HeaderGroupName: "Header and Button Style",
    QuestionGroupName: "Question Style",
    AnswerGroupName: "Answer Style",
    ListNameLabel: "List Name",
    TitleLabel: "Title",
    MaxItemsPerPageLabel: "Max number of items per page",
    MaxItemsToFetchFromTheListLabel:
      "Max number of items to fetch from the list",
    HeaderBackgroundColorLabel: "Header Background Colour",
    HeaderTextColorLabel: "Header Text Colour",
    QuestionBackgroundColorLabel: "Question Background Colour",
    QuestionTextColorLabel: "Question Text Colour",
    AnswerBackgroundColorLabel: "Answer Background Colour",
    AnswerTextColorLabel: "Answer Text Colour",
    ResetStyleButtonText: "Reset to Default",
    isSearchAbleText: "Display Search Box",
    queryFilterPanelStrings: {
      filtersLabel: "Filters",
      addFilterLabel: "Add filter", 
      loadingFieldsLabel: "Loading fields from specified list...",
      loadingFieldsErrorLabel: "An error occured while loading fields : {0}",
      queryFilterStrings: {
        fieldLabel: "Field",
        fieldSelectLabel: "Select a field...",
        operatorLabel: "Operator",
        operatorEqualLabel: 'Equals',
        operatorNotEqualLabel: 'Does not equal',
        operatorGreaterLabel: 'Is greater than',
        operatorGreaterEqualLabel: 'Is greater or equal to',
        operatorLessLabel: 'Is less than',
        operatorLessEqualLabel: 'Is less or equal to',
        operatorContainsLabel: 'Contains',
        operatorBeginsWithLabel: 'Begins with',
        operatorContainsAnyLabel: 'Contains Any',
        operatorContainsAllLabel: 'Contains All',
        operatorIsNullLabel: 'Is Null',
        operatorIsNotNullLabel: 'Is Not Null',
        valueLabel: 'Value',
        andLabel: 'And',
        orLabel: 'Or',
        peoplePickerSuggestionHeader: 'Suggested People',
        peoplePickerNoResults: 'No results found',
        peoplePickerLoading: 'Loading users',
        peoplePickerMe: 'Me',
        taxonomyPickerSuggestionHeader: 'Suggested Terms',
        taxonomyPickerNoResults: 'No results found',
        taxonomyPickerLoading: 'Loading terms',
        datePickerLocale: 'en',
        datePickerFormat: 'MMM Do YYYY, hh:mm a',
        datePickerExpressionError: 'Expression must respect the following format : [Today] or [Today] +/- [digit]',
        datePickerDatePlaceholder: 'Select a date...',
        datePickerExpressionPlaceholder: 'Or enter a valid expression...',
        datePickerIncludeTime: 'Include time in query',
        datePickerStrings: {
          months: [ 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December' ],
          shortMonths: [ 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec' ],
          days: [ 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday' ],
          shortDays: [ 'S', 'M', 'T', 'W', 'T', 'F', 'S' ],
          goToToday: 'Go to today'
        }
      }
    }
  };
});

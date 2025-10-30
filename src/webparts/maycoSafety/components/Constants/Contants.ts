export enum ControlType {
    number='Number',
    string='String',
    mobileNumber='MobileNumber',
    email='Email',
    people='PeoplePicker',
    date='DatePicker',
    DateTime='DateTimePicker',
    compareDates='CompareDates',
    compareNum='CompareNumbers',
    lessthanTodayDate='LessThanTodayDate',
    greaterthanTodayDate='GreaterThanTodayDate',
    reactSelect='ReactSelect',
    radio='Radio',
    array='Array',
    limitedNumber="LimitedNumber"
  }
  export enum ActionStatus{
    Draft='saved successfully',
    Submitted='submitted successfully',
    Updated='updated successfully',
    Approved='approved successfully',
    Rejected='rejected successfully',
    Error='Sorry! something went wrong',
    Export='exported successfully'
  }
  export enum FormStatus{
    Draft='Saved',
    Submitted='Submitted',
    Approved='Approved',
    Rejected='Rejected',
    Proceeded='Proceeded',
  }
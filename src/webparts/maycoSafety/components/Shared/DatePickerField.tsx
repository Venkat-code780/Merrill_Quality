import * as React from "react";
import { useState } from "react";
import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";



const DatePickercontrol = (props: any) => {
  var [selectedDay, setDate] = useState(null);
  let selectedDate = props.selectedDate != null ? props.selectedDate : null;
  let startDate = props.startDate;
  let endDate = props.endDate;
  // let selDate=null;
  if (selectedDate != null) {
    selectedDay = selectedDate;
  }
  else {
    selectedDay = null;
  }

  if (props.isDisabled) {
    setTimeout(() => {
      var DatePickers = document.getElementsByClassName("DatePicker");
      for (var i = 0; i < DatePickers.length; i++) {
        (DatePickers[i] as HTMLInputElement).disabled = true;
      }
    }, 1000);
  } else {
    setTimeout(() => {
      var DatePickers = document.getElementsByClassName("DatePicker");
      for (var i = 0; i < DatePickers.length; i++) {
        (DatePickers[i] as HTMLInputElement).disabled = false;
      }
    }, 1000);
  }

  function handlechangeevent(seldate: any) {
    setDate(seldate);
    props.onDatechange([seldate, props.id, props.name]);
  }
  return (
    <DatePicker
      selected={selectedDay}
      title={props.title}
      dateFormat={props.showTime ? props.TimeFormat : 'MM/dd/yyyy'}
      // timeFormat="hh:mm aa" // Forces 12-hour format
      // timeIntervals={1}    // Optional, for nicer time stepping
      // timeCaption="Time"    // Optional label for the time selector
      // showTimeSelect        // Enables the time dropdown selectorF
      showBorder={true}
      onChange={handlechangeevent}
      highlightDates={[props.highlightDate]}
      placeholderText={props.placeholder}
      className="form-control DatePicker"
      id={props.id}
      readOnly={props.readOnly || false}
      disabled={props.isDisabled || false}
      showIcon
      showMonthDropdown
      showYearDropdown
      toggleCalendarOnIconClick
      minDate={[null, undefined, ''].includes(startDate) ? undefined : startDate}
      maxDate={[null, undefined, ''].includes(endDate) ? undefined : endDate}
      showTimeInput={props.showTime ?? false}
    // ref={endDate.ref || endDate
    />
  );
};

export default DatePickercontrol;
import * as React from "react";
import { useState, useEffect } from "react";
import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";

interface DatePickerControlProps {
  id?: string;
  name?: string;
  title?: string;
  selectedDate?: string;
  startDate?: Date | null;
  endDate?: Date | null;
  isDisabled?: boolean;
  readOnly?: boolean;
  placeholder?: string;
  showTime?: boolean;
  highlightDate?: Date | null;
  showIcon?: boolean;
  onDatechange?: (args: [Date | null, string | undefined, string | undefined]) => void;
}

const DatePickercontrol = React.forwardRef<any, DatePickerControlProps>((props, ref) => {
  const [selectedDay, setSelectedDay] = useState<Date | null>(
    props.selectedDate ? new Date(props.selectedDate) : null
  );

  useEffect(() => {
    if (props.selectedDate) {
      setSelectedDay(new Date(props.selectedDate));
    } else {
      setSelectedDay(null);
    }
  }, [props.selectedDate]);

  // Handle date selection from the picker
  const handleChange = (selDate: Date | null) => {
    setSelectedDay(selDate);
    props.onDatechange?.([selDate, props.id, props.name]);
  };

  // Handle manual input (MM/DD/YYYY)
  const handleRawChange = (e: React.ChangeEvent<HTMLInputElement>) =>{
    const inputValue = e.target.value;
    const dateRegex = /^(0[1-9]|1[0-2])\/(0[1-9]|[12][0-9]|3[01])\/\d{4}$/;

    if (dateRegex.test(inputValue)) {
      const parsedDate = new Date(inputValue);
      let now = new Date();
      const isAfterStart =
        !props.startDate || parsedDate >= new Date(props.startDate.setHours(0, 0, 0));
      const isBeforeEnd =
        !props.endDate || parsedDate <= new Date(props.endDate.setHours(now.getHours(), now.getMinutes(), now.getSeconds()));

      if (!isNaN(parsedDate.getTime()) && isAfterStart && isBeforeEnd) {
        props.onDatechange?.([parsedDate, props.id, props.name]);
      }
    } else {
      props.onDatechange?.([null, props.id, props.name]);
    }
  };

  // Handle “Now” button click

  // const handleNowClick = () => {
  //   const now = new Date();
  //   setSelectedDay(now);
  //   props.onDatechange?.([now, props.id, props.name]);
  //   datePickerRef.current?.setOpen(false);
  // };

  // ✅ Choose date format dynamically
  const dateFormat = props.showTime ? "MM/dd/yyyy hh:mm aa" : "MM/dd/yyyy";

  return (
    <DatePicker
      ref={ref}
      selected={selectedDay}
      onChange={handleChange}
      onChangeRaw={handleRawChange}
      placeholderText={props.placeholder ?? (props.showTime ? "MM/DD/YYYY hh:mm AM" : "MM/DD/YYYY")}
      dateFormat={dateFormat}
      highlightDates={props.highlightDate ? [props.highlightDate] : []}
      className="form-control DatePicker"
      id={props.id}
      name={props.name}
      readOnly={props.readOnly || false}
      disabled={props.isDisabled || false}
      showIcon={props.showIcon ?? true}
      showMonthDropdown
      showYearDropdown
      toggleCalendarOnIconClick
      minDate={props.startDate || undefined}
      maxDate={props.endDate || undefined}
      tabIndex={0}
      title={props.title}
      showTimeInput={props.showTime ?? false}
      // Below props are for Time Drpdown format
      //Add this block to inject "Now" button inside popup
      // showTimeSelect={props.showTime ?? false}
      // timeFormat="hh:mm aa"
      // timeIntervals={1} // 1 minute gap
      // timeCaption="Time"
      // calendarContainer={({
      //   className,
      //   children,
      // }: {
      //   className?: string;
      //   children?: React.ReactNode;
      // }) => (
      //   <div className={className}>
      //     {children}
      //     {props.showTime && !props.isDisabled && (
      //       <div style={{padding: "8px" }}>
      //         <button
      //           type="button"
      //           onClick={handleNowClick}
      //           style={{
      //             backgroundColor: "#1976d2",
      //             color: "white",
      //             border: "none",
      //             borderRadius: "4px",
      //             padding: "4px 10px",
      //             fontSize: "0.8rem",
      //             cursor: "pointer",
      //             minWidth: '20px'
      //           }}
      //         >
      //           Now
      //         </button>
      //       </div>
      //     )}
      //   </div>
      // )}
    />
  );
});

export default DatePickercontrol;

import * as React from "react";
import { useState, useEffect } from "react";
// import TextField from "@mui/material/TextField";
import { AdapterDayjs } from "@mui/x-date-pickers/AdapterDayjs";
import { LocalizationProvider } from "@mui/x-date-pickers/LocalizationProvider";
import { DateTimePicker } from "@mui/x-date-pickers/DateTimePicker";
import * as dayjs from "dayjs";
//import { TextField } from "@mui/material";

// import IconButton from "@mui/material/IconButton";
// import InputAdornment from "@mui/material/InputAdornment";
// import ClearIcon from "@mui/icons-material/Clear";
type Dayjs = typeof dayjs extends { default: infer D } ? D : ReturnType<typeof dayjs>;

interface DatePickerControlProps {
  id?: string;
  name?: string;
  title?: string;
  selectedDate?: Date | null;
  startDate?: Date | null;
  endDate?: Date | null;
  isDisabled?: boolean;
  readOnly?: boolean;
  placeholder?: string;
  showTime?: boolean;
  onDatechange?: (args: [Date | null, string | undefined, string | undefined]) => void;
}

const DateTimePickercontrol: React.FC<DatePickerControlProps> = (props) => {
  const [selectedDate, setSelectedDate] = useState<Dayjs | null>(
    props.selectedDate ? dayjs(props.selectedDate) : null
  );
  //const [open, setOpen] = useState(false);
  useEffect(() => {
    setSelectedDate(props.selectedDate ? dayjs(props.selectedDate) : null);
  }, [props.selectedDate]);


  const handleChange = (newValue: Dayjs | null) => {
    if (newValue === null || dayjs(newValue).isValid()) {
      //const minDate = props.startDate ? dayjs(props.startDate) : null;
      //const maxDate = props.endDate ? dayjs(props.endDate) : null;

      const parsedDate = newValue ? newValue.toDate() : null;
      let now = new Date();
      const isAfterStart = !(props.startDate) || (parsedDate && parsedDate >= new Date(props.startDate.setHours(now.getHours(), now.getMinutes(), now.getSeconds())));
      const isBeforeEnd = !(props.endDate) || (parsedDate && parsedDate <= new Date(props.endDate.setHours(now.getHours(), now.getMinutes(), now.getSeconds())));
      if (parsedDate && !isNaN(parsedDate.getTime()) && isAfterStart && isBeforeEnd) {
        setSelectedDate(newValue);
        //setOpen(false)
        props.onDatechange?.([
          newValue ? newValue.toDate() : null,
          props.id,
          props.name,
        ]);
      } else if (newValue === null) {
        setSelectedDate(null);
        //setOpen(false)
        props.onDatechange?.([null, props.id, props.name]);
      }
    }
  };
  // const handleClear = () => {
  //   setSelectedDate(null);
  //   props.onDatechange?.([null, props.id, props.name]);
  // };

  return (
    <LocalizationProvider dateAdapter={AdapterDayjs}>
      <DateTimePicker
        label={props.title}
        value={selectedDate}
        onChange={handleChange}
        disabled={props.isDisabled ?? false}
        readOnly={props.readOnly ?? false}
        minDateTime={props.startDate ? dayjs(props.startDate) : undefined}
        maxDateTime={props.endDate ? dayjs(props.endDate) : undefined}
        ampm // enables AM/PM dropdown
        timeSteps={{ hours: 1, minutes: 1 }}
        //open={open}
        //onOpen={() => setOpen(true)}
        //onClose={() => setOpen(false)}
        //enableAccessibleFieldDOMStructure={false}
        localeText={{ todayButtonLabel: 'Now' }}
        slotProps={{
          textField: {
            //component: TextField,
            id: props.id,
            name: props.name,
            placeholder: props.placeholder ?? "MM/DD/YYYY hh:mm A",
            fullWidth: true,
            required: false,
            size: "small",
            label: '',
            title: selectedDate && dayjs(selectedDate).isValid()
              ? dayjs(selectedDate).format('MM/DD/YYYY hh:mm A')
              : '',
            //onFocus: () => setOpen(true),
            // InputProps: {
            //   endAdornment: (
            //     <>
            //       {selectedDate && (
            //         <InputAdornment position="end">
            //           <IconButton onClick={handleClear} size="small">
            //             <ClearIcon fontSize="small" />
            //           </IconButton>
            //         </InputAdornment>
            //       )}
            //     </>
            //   ),
            // },
          },
          actionBar: {
            actions: ['today', 'accept', 'cancel'], // 'today' displays as 'Now' in the DateTimePicker
          },
        }}
        className="form-control DatePicker"
      />
    </LocalizationProvider>
  );
};

export default DateTimePickercontrol;



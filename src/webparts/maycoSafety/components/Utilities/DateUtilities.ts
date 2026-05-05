import { format } from "date-fns";
class DateUtilities {
  public static getDateMMDDYYYY(givenDate: any) {
    let date = new Date(givenDate);
    return (date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1) + "/" + (date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()) + "/" + date.getFullYear();
  }
  public static getDateMMDDYYYYTime(givenDate: any) {
    let date = new Date(givenDate);
    return (date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1) + "/" + (date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()) + "/" + date.getFullYear() + " " + date.getHours() + ":" + date.getMinutes();
  }
  public static getDateMMDDYYYYTimes(givenDate: any) {
    let date = new Date(givenDate);

    // Get hours, minutes, and AM/PM
    let hours = date.getHours();
    let minutes: number | string = date.getMinutes();  // Allow both number and string for minutes
    let ampm = hours >= 12 ? "PM" : "AM";

    // Convert hours from 24-hour format to 12-hour format
    hours = hours % 12;
    hours = hours ? hours : 12; // the hour '0' should be '12'

    // Ensure minutes are formatted as a 2-digit string
    minutes = minutes < 10 ? "0" + minutes : minutes.toString(); // Convert to string

    // Return the formatted date with AM/PM
    return (date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1) + "/" +
      (date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()) + "/" + date.getFullYear() +
      " " + hours + ":" + minutes + " " + ampm;
  }


  public static getDateYYYYMMDDForSorting(givenDate: any) {
    let date = new Date(givenDate);
    return `${date.getFullYear()}${date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1}${date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()}`;
  }
  public static getDateYYYYMMDDHHMMForSorting(givenDate: string) {
    let date = new Date(givenDate);
    let formatedDate = format(date,'yyyy-MM-dd HH:mm');
    // let YYYYMMDDHHMM = formatedDate.replace(/-/g, '').replace(/\//g, '').replace(/:/g, '').replace(/ /g, ''); 
    let YYYYMMDDHHMM = formatedDate.replace(/[-/: ]/g, ''); // short form of above line

    return YYYYMMDDHHMM; // YYYYMMDDHHMMSS
  }
  public static addBrowserwrtServer(date: Date, webTimeZoneData: any) {
    var utcOffsetMinutes = date.getTimezoneOffset();
    var newDate = new Date(date.getTime());
    newDate.setTime(newDate.getTime() + ((webTimeZoneData.Bias - utcOffsetMinutes + webTimeZoneData.DaylightBias) * 60 * 1000));
    return newDate;
  }

  public static removeBrowserwrtServer(date: Date, webTimeZoneData: any) {
    var utcOffsetMinutes = date.getTimezoneOffset();
    var newDate = new Date(date.getTime());
    newDate.setTime(newDate.getTime() - ((webTimeZoneData.Bias - utcOffsetMinutes + webTimeZoneData.DaylightBias) * 60 * 1000));
    return newDate;
  }
  public static getDateYYYYMMDD(givenDate: any) {
  if (!givenDate) return "";

  const date = new Date(givenDate);

  const yyyy = date.getFullYear();
  const mm = ("0" + (date.getMonth() + 1)).slice(-2);
  const dd = ("0" + date.getDate()).slice(-2);

  return `${yyyy}-${mm}-${dd}`;
}
public static addDays(givenDate: any, days: number) {
  const date = new Date(givenDate);
  date.setDate(date.getDate() + days);
  return date;
}
public static formatYYYYMMDD(date: Date) {
  if (!date) return "";

  const yyyy = date.getFullYear();
  const mm = ("0" + (date.getMonth() + 1)).slice(-2);
  const dd = ("0" + date.getDate()).slice(-2);

  return `${yyyy}-${mm}-${dd}`;
}


}
export default DateUtilities;
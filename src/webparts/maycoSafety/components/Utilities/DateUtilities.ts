class DateUtilities {
    public static getDateMMDDYYYY(givenDate:any)
    {
        let date=new Date(givenDate);
          return (date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1) + "/" + (date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()) + "/" + date.getFullYear();
    }
    public static getDateMMDDYYYYTime(givenDate:any)
    {
        let date=new Date(givenDate);
          return (date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1) + "/" + (date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()) + "/" + date.getFullYear()+ " " + date.getHours() + ":" + date.getMinutes();
    }
    public static getDateYYYYMMDDForSorting(givenDate: any)
    {
        let date=new Date(givenDate);
          return `${date.getFullYear()}${date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1}${date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()}` ;
    }
    
    public static addBrowserwrtServer(date:Date, webTimeZoneData:any) {
        var utcOffsetMinutes = date.getTimezoneOffset();
        var newDate = new Date(date.getTime());
        newDate.setTime(newDate.getTime() + ((webTimeZoneData.Bias - utcOffsetMinutes + webTimeZoneData.DaylightBias) * 60 * 1000));
        return newDate;
    }
 }
 export default DateUtilities;
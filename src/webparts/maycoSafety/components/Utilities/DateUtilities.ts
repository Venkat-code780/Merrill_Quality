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

    public static getFriendlyDate(givenDate: any): string {
        if (!givenDate) return "";
        let date = new Date(givenDate);
        const now = new Date();

        const diffMs = now.getTime() - date.getTime();
        const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

        if (diffDays === 0) {
            return "Today";
        } else if (diffDays === 1) {
            return "Yesterday";
        } else if (diffDays <= 7) {
            return `${diffDays} days ago`;
        } else {
            return date.toLocaleDateString("en-US", {
                year: "numeric",
                month: "long",
                day: "numeric"
            }); // e.g., "August 13, 2020"
        }
    }

     public static removeBrowserwrtServer(date: Date, webTimeZoneData: any) {
    var utcOffsetMinutes = date.getTimezoneOffset();
    var newDate = new Date(date.getTime());
    newDate.setTime(newDate.getTime() - ((webTimeZoneData.Bias - utcOffsetMinutes + webTimeZoneData.DaylightBias) * 60 * 1000));
    return newDate;
  }
 }
 export default DateUtilities;
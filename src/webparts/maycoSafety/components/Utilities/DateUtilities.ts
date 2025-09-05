class DateUtilities {
    public static getDateMMDDYYYY(givenDate:any)
    {
        let date=new Date(givenDate);
          return (date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1) + "/" + (date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()) + "/" + date.getFullYear();
    }
    public static getDateYYYYMMDDForSorting(givenDate: any)
    {
        let date=new Date(givenDate);
          return `${date.getFullYear()}${date.getMonth() < 9 ? "0" + (date.getMonth() + 1) : date.getMonth() + 1}${date.getDate() <= 9 ? "0" + date.getDate() : date.getDate()}` ;
    }
 }
 export default DateUtilities;
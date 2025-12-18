import { Guid } from "@microsoft/sp-core-library";

/**
 * @param {Date | string} date date to which the no of days should be added.
 * @param {number} noOfDays no of days you want the future date to be.
 * @returns {Date} the future date
*/
export const getFutureDate = (date: Date | string, noOfDays: number): Date => {
    return new Date(new Date(date).getTime() + (noOfDays * 24 * 60 * 60 * 1000));
};

/**
 * @param {Date | string} date date to which the no of days should be subtracted.
 * @param {number} noOfDays no of days you want the past date to be.
 * @returns {Date} the past date
*/
export const getPastDate = (date: Date | string, noOfDays: number): Date => {
    return new Date(new Date(date).getTime() - (noOfDays * 24 * 60 * 60 * 1000));
};


/**
 * @param {string | number | Date} date the date to be formated
 * @returns {string} formatted date
 */
export const formatDate = (date: string | number | Date, numbered?: boolean): string => {
    if (numbered) {
        const day = new Date(date).getDate().toString().padStart(2, '0');   // Adds leading zero to single-digit days
        const month = (new Date(date).getMonth() + 1).toString().padStart(2, '0'); // Adds leading zero to single-digit months
        const year = new Date(date).getFullYear();

        // Return formatted date as "dd-mm-yyyy"
        return `${day}-${month}-${year}`;
    } else {
        const formattedDate = new Date(date).toDateString().split(" ");
        return `${formattedDate[2]}-${formattedDate[1]}-${formattedDate[3]}`;
    }
};


/**
 * @param {Date | string} date
 * @returns {Date} date without timezone
 */


/**
 * @param {Date|string|number} date the date to be converted to DayPilot date 
 * @returns the DayPilot Date that snips to the cells
 */

/**
 * @param {string} html the html in string format
 * @returns the content between html tags
 */
export const getHTMLContent = (html: string) => {
    const regex = /<.*?>(.*?)<\/.*?>/;
    const match = html.match(regex);
    if (match) {
        return match[1];
    }
    return null;
};

/**
 * returns the date difference between two dates (including the dates)
 * @param date1: Start Date
 * @param date2: End Date
 * @param includeWeekends: The inlude the weekens or not
 * @param includeBothDates: The inlude the start and end in the difference
 * @returns the difference
 */

/**
 * 
 * @param str string to be sanitized
 * Allows only a-z A-Z 0-9 and _
 * @returns sanitized string
 */
export const santizeInput = (str: string, allowedComma?: boolean) => {
    let regex: RegExp;

    if (allowedComma) {
        regex = /[^a-zA-Z0-9\-_\, ]/g;
    } else {
        regex = /[^a-zA-Z0-9\-_ ]/g;
    }

    if (str) {
        return str.replace(regex, "");
    }
};

/**
 * @param {Date | string} date
 * @returns {Date} date without timezone
 */
export const getOriginalDate = (date: string): Date => {
    if (!date) return null;

    if (date.includes("/")) {
        const [month, day, year] = date.split("/");
        return new Date(Number(year), Number(month) - 1, Number(day));
    } else {
        const [day, month, year] = date.split("-");
        return new Date(Number(year), Number(month) - 1, Number(day));
    }
};

export const getUUID = () => Guid.newGuid().toString();


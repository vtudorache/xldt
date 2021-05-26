#ifndef __XLDT_DOC_H__
#define __XLDT_DOC_H__

#ifndef PyDoc_STRVAR
#error The Python header was not included or it's too old.
#endif

PyDoc_STRVAR(xldt__doc__,
"Excel compatible date and time helper functions for Python.\n\n\
This extension handles date and time using serial numbers, just like\n\
Microsoft Excel. The serial numbers are identical between this module\n\
and Excel for dates starting with 1900-03-01 in the 1900 date system.");

PyDoc_STRVAR(xldt_date__doc__,
"date(year: float, month: float, day: float) -> float\n\n\
Return the value corresponding to the given date. The month and the\n\
day are optional, defaulting to 1.");

PyDoc_STRVAR(xldt_day__doc__,
"day(value: float) -> int\n\n\
Return the day (1 - 31) of month of the date corresponding to the\n\
given value.");

PyDoc_STRVAR(xldt_days__doc__,
"days(start_date: float, end_date: float) -> int\n\n\
Calculate the number of days between two dates.");

PyDoc_STRVAR(xldt_hour__doc__,
"hour(value: float) -> int\n\n\
Return the hour (0 - 23) corresponding to the given value.");

PyDoc_STRVAR(xldt_isweekend__doc__,
"isweekend(value: float, result_type: int) -> bool\n\n\
Return True if the date corresponding to the given value is a weekend.\n\
The optional result_type argument controls which of the days in a week\n\
are considered weekend. It can take on any of the constants WE_SAT_SUN,\n\
WE_SUN_MON, WE_MON_TUE, WE_TUE_WED, WE_WED_THU, WE_THU_FRI, WE_FRI_SAT,\n\
WE_SUN, WE_MON, WE_TUE, WE_WED, WE_THU, WE_FRI, WE_SAT.\n\
The result_type can also be a seven characters string, each character\n\
can be 1 (weekend) or 0 (workday).\n\
The result type defaults to WE_SAT_SUN.\n\
The values of the constants are identical to those used by Excel with\n\
WORKDAY.INTL and could be used to implement equivalent functions in\n\
Python.");

PyDoc_STRVAR(xldt_isoweek__doc__,
"isoweek(serial: float) -> int\n\n\
Return the ISO week number of the year for a given date.\n\
It behaves like the Excel ISOWEEKNUM function.");

PyDoc_STRVAR(xldt_minute__doc__,
"minute(value: float) -> int\n\n\
Return the minute (0 - 59) corresponding to the given value.");

PyDoc_STRVAR(xldt_month__doc__,
"month(value: float) -> int\n\n\
Return the month (1 - 12) of the date corresponding to the given value.");

PyDoc_STRVAR(xldt_months__doc__,
"months(start_date: float, end_date: float) -> int\n\n\
Calculate the number of full months between the dates corresponding to\n\
the given values.");

PyDoc_STRVAR(xldt_now__doc__,
"now() -> float\n\n\
Return the value corresponding to the current local date and time.");

PyDoc_STRVAR(xldt_second__doc__,
"second(value: float) -> int\n\n\
Return the second (0 - 59) corresponding to the given value.");

PyDoc_STRVAR(xldt_time__doc__,
"time(hour: float, minute: float, second: float) -> float\n\n\
Return the value corresponding to the given time of the day. The \n\
minute and the second are optional, defaulting to 0.");

PyDoc_STRVAR(xldt_today__doc__,
"today() -> int\n\n\
Return the value corresponding to the current date (without time).");

PyDoc_STRVAR(xldt_week__doc__,
"week(serial: float, result_type: int) -> int\n\n\
Return the week number corresponding to the given serial number. The\n\
result_type argument is optional, defaulting to 1.\n\
It behaves like the Excel WEEKNUM function.");

PyDoc_STRVAR(xldt_weekday__doc__,
"weekday(value: float, result_type: float) -> int\n\n\
Return the weekday corresponding to the given value. The result_type\n\
argument is optional, defaulting to 1.");

PyDoc_STRVAR(xldt_year__doc__,
"year(value: float) -> int\n\n\
Return the year of the date corresponding to the given value.");

PyDoc_STRVAR(xldt_years__doc__,
"years(start_date: float, end_date: float) -> int\n\n\
Calculate the number of full years between the dates corresponding to\n\
the given values.");

#endif

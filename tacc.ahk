TTAC(){
; data := [
;     ["06/12/2024", "48202: Add OTP #, Prepared For, and Description fields to cXML records", 6], 
;     ["06/12/2024", "48688: Adjust DM Wizard Letter product verbiage", 2], 
; ]
    ;;;DATA;;;
data := [
    ["08/26/2024", "51005: Add Total Number of Quotes per product to report", 4], 
    ["08/20/2024", "51005: Add Total Number of Quotes per product to report", 5], 
    ["08/19/2024", "51005: Add Total Number of Quotes per product to report", 8], 
]
;;;END;;;

    long_sleep := 2000
    short_sleep := 500

    for index, item in data
    {
        Run "https://tacc.tula.co/timesheet/working/704"
        Sleep long_sleep ; Wait for 3 seconds to allow the page to load
        ; open browser console
        Send "{F12}"
        Sleep long_sleep

        date := item[1]
        desc := item[2]

        if (item[3] != "") {
            hours := item[3]
        } else {
            hours := 8
        }
        Send "$('{#}Date').val('" date "');"
        Send "{Enter}"
        Sleep short_sleep

        Send "$('{#}Description').val('" desc "');"
        Send "{Enter}"
        Sleep short_sleep

        Send "$('{#}TotalHourId').val('" hours "');"
        Send "{Enter}"
        Sleep short_sleep

        Send "$('{#}form > form > button').click();"
        Send "{Enter}"
        Sleep short_sleep
        ; close tab with ctrl + w
        Send "{CtrlDown}w{CtrlUp}"
    }

}
!t::TTAC()

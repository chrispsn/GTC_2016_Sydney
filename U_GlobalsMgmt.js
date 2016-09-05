'use strict'

var U_GlobalsMgmt = {
       
    temp_disable_screenupdating: function(excel_app, func) {
        var original_setting = excel_app.ScreenUpdating
        excel_app.ScreenUpdating = false;
        var output = func();
        excel_app.ScreenUpdating = original_setting;
        return output;
    },

    temp_disable_displayalerts: function(excel_app, func) {
        var original_setting = excel_app.DisplayAlerts
        excel_app.DisplayAlerts = false;
        var output = func();
        excel_app.DisplayAlerts = original_setting;
        return output;
    }

}

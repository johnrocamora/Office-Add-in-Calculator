/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/


/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $("#buttonMrc").click(function (event) { getFromMemory(); });
            $("#buttonAddToMemory").click(function (event) { inputFromDisplay(); });
            $("#buttonClearMemory").click(function (event) { setValue(emptyString); });
            $("#buttonOne").click(function (event) { appendValue('1'); });
            $("#buttonTwo").click(function (event) { appendValue('2'); });
            $("#buttonThree").click(function (event) { appendValue('3'); });
            $("#buttonFour").click(function (event) { appendValue('4'); });
            $("#buttonFive").click(function (event) { appendValue('5'); });
            $("#buttonSix").click(function (event) { appendValue('6'); });
            $("#buttonSeven").click(function (event) { appendValue('7'); });
            $("#buttonEight").click(function (event) { appendValue('8'); });
            $("#buttonNine").click(function (event) { appendValue('9'); });
            $("#buttonZero").click(function (event) { appendValue('0'); });
            $("#buttonDot").click(function (event) { appendValue('.'); });
            $("#buttonClear").click(function (event) { setValue(emptyString); });
            $("#buttonMultiply").click(function (event) { appendValue('*'); });
            $("#buttonDivide").click(function (event) { appendValue('/'); });
            $("#buttonAdd").click(function (event) { appendValue('+'); });
            $("#buttonSubtract").click(function (event) { appendValue('-'); });
            $("#buttonEquals").click(function (event) { calculate(); });
            $('#set-selected-data').click(function (event) { setSelectedDataFromDisplay(); });
            $('#get-selected-data').click(function (event) { getDataFromSelection(); });
            setValue(emptyString);
        });
    };

    var emptyString = ''; //global string
    var inMemory = 'idOfSetting';
    var inputMemory = '';

    // Insert the data from the display box into the active slide
    function setSelectedDataFromDisplay() {
        Office.context.document.setSelectedDataAsync(document.getElementById("d").value,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    app.showNotification('Error:', result.error.message);

                }
            }
        );
    }

    //Insert the data from the slide into the display box.
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var valueOfResult = result.value;
                    setValue(valueOfResult);
                }
                else
                    app.showNotification('Error:', result.error.message);
            });
    }

    //Sets the value of the custom settings in memory to be used later. 
    function setValue(val) {
        Office.context.document.settings.set(inMemory, val);
        document.getElementById("d").value = val;
    }

    //Appends a string to the settings value.
    function appendValue(val) {
        var newValue = Office.context.document.settings.get(inMemory) + val;
        Office.context.document.settings.set(inMemory, newValue);
        document.getElementById("d").value += val;
    }

    //Save a number from display.
    function inputFromDisplay() {
        inputMemory = document.getElementById("d").value; //how do we store the string
        app.showNotification('Number saved: ', inputMemory);
    }

    //Get the Setting value that was stored in memory.
    function getFromMemory() {
        appendValue(inputMemory);
    }

    //Calculate what's in the Settings object, and display it.
    function calculate() {
        try {
            var calculatedValue = eval(Office.context.document.settings.get(inMemory));
            document.getElementById("d").value = calculatedValue;

            app.showNotification('Calculated Value: ', calculatedValue.toString());
        }
        catch (err) {
            app.showNotification("Error: ", "invalid characters");
        }
    }
})();

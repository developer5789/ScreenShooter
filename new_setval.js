function setval(number, datefrom, dateto){
    if(
        document.getElementById('combobox-1032-inputEl')
    ){
        if(document.getElementById('combobox-1032-inputEl').value != 'Интервал'){
            console.log('done')
            document.getElementById('combobox-1032-inputEl').value = 'Интервал';
            $('#datefield-1033-inputEl').show();
            $('#datefield-1034-inputEl').show();
            $("#button-1165-btnIconEl").click();
            $("#ext-183-btnIconEl").click();

            $('tabpanel-1321').hide()
            $('panel-1326').hide()


        }
        document.getElementById('textfield-1123-inputEl').value = number;
        document.getElementById('datefield-1033-inputEl').value = datefrom;
        document.getElementById('datefield-1034-inputEl').value = dateto;
        $("#button-1154-btnIconEl").click();

        setTimeout("$('.x-grid-item:first-child .x-grid-cell:nth-child(2)').click()",1000);
    }
    else{
        alert("Внимание, компоненты интерфейса изменились! Попробуйте перезапустить браузер!");
    }
};

$('#datefield-1033-inputEl').show();
$('#datefield-1034-inputEl').show();
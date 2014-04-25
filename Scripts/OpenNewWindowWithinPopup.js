(function ($) {
    $.OpenNewWindowPopUp = function (title, src, width, height) {
        //Destroy if exist
        $('#dv_move').remove();
        //create hte popup html
        var html = '<div class="main" id="dv_move" style="width:' + width + 'px; height:' + height + 'px;">';
        html += '  <div class="title">' + title + '</div>';
        html += ' <div id="dv_no_move">';
        html += '<div id="dv_load"><img src="../images/HourlySales/circular.gif"/></div>';
        html += ' <iframe id="url" scrolling="auto" src="' + src + '"  style="border:none;" width="100%" height="100%"></iframe>';
        html += ' </div>';
        html += ' </div>';

        //add to body
        $('<div></div>').prependTo('body').attr('id', 'overlay'); // add overlay div to disable the parent page
        $('body').append(html);


        $("#dv_no_move").mousedown(function () {
            return false;
        });

        $("#title").mousedown(function () {
            return false;
        });

        setTimeout("$('#dv_load').hide();", 500);
    };
})(jQuery);

//close popup
function CloseDialog() {
    $('#overlay').fadeOut('slow');
    $('#dv_move').fadeOut('slow');
    setTimeout("$('#dv_move').remove();", 1000);

    //call Refresh(); if we need to reload the parent page on its closing
    // parent.Refresh();
}
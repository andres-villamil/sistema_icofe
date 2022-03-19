function fnc_show_1() {
    $('#contenido_p3mb').show();
}
function fnc_hide_1() {
    $('#contenido_p3mb').hide();
}
function fnc_show_2() {
    $('#contenido_p7mc').show();
}
function fnc_hide_2() {
    $('#contenido_p7mc').hide();
}
function fnc_show_3() {
    $('#contenido_p10mc').show();
}
function fnc_hide_3() {
    $('#contenido_p10mc').hide();
}
function fnc_show_4() {
    $('#contenido_p8mf').show();
}
function fnc_hide_4() {
    $('#contenido_p8mf').hide();
}
function fnc_show_5() {
    $('#contenido_p9mf').show();
}
function fnc_hide_5() {
    $('#contenido_p9mf').hide();
}
function fnc_hide_sidebar() {
    $('#navbarOpen').hide();
    $('#navbarClose').show();
    $("#content").addClass("col-lg-11");
    $("#content").removeClass("col-lg-10");
}
function fnc_show_sidebar() {
    $('#navbarOpen').show();
    $('#navbarClose').hide();
    $("#content").removeClass("col-lg-11");
    $("#content").addClass("col-lg-10");
}
$(function () {
    $('#avatar').click(function () {
        $('#popoverAvatar').toggle();
    });
});
function searchOpen(){    
    $("#btn-Small").hide();
    $('.buscador-big').fadeIn('slow');
    
}

function searchClose(){
    $("#btn-Small").fadeIn('slow');
    $('.buscador-big').hide();
   
    
}




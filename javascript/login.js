$( document ).ready(function(){
    $('.mdi-close').on("click",function(){
        $('#pro-banner').slideUp();
    })
    $('#registerAcc').on("click",function(){		
        $('#visuSoftware').modal('show');
    })
})
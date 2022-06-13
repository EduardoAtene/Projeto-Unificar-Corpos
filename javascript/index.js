$(document).ready(function() {
    $(document).on("click", "a", function(e){

        var href = $(this).attr("href");	

        if((href.includes("*"))){
            let route = location.pathname.split("/");
            console.log(location);
            var url = location.protocol + '//' + location.host + "/" +  href.split("*")[1];
            //var url = location.protocol + '//' + location.host + route[1] + route[2] + "/"  + href.split("*")[1];   

                                action = function(){ window.location.replace(url);
                                }
                                setTimeout(action, 1);
        }		
        
        //return true;
    });
});
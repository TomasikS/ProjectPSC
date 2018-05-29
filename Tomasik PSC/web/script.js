$(function(){

    $('#findPsc').on('click', function(){
        var $town = $('#input1').val();


        $.ajax({
            url:"http://localhost:9090/posta/city/"+ $town,
            data: { format: 'json'},
            dataType: "json",
            error: function(){
                $('#pscRes').html('Not found');
            },
            success : function(data){
            $("#cityRes").empty();
            var i=0;
            var ress=$("#cityRes");
            var nPsc=$("<div>"+data[i]+"</div>");

            for(i=0; i<data.length;i++)
            {
            ress.append("<div>"+data[i]+"</div> ");
            }
            },
            type: 'GET'

        });
    });

    $('#findCity').on('click', function(){
        var $psc = $('#input2').val();


        $.ajax({
            url:"http://localhost:9090/posta/psc/"+ $psc,
            data: { format: 'json'},
            dataType: "json",
            error: function(){
                $('#cityRes').html('Not found');
            },
            success : function(data){
            $("#cityRes").empty();
             var i=0;
             var ress=$("#cityRes");

            for(i=0; i<data.length;i++)
                ress.append("<div>"+data[i]+"</div> ");
            },
            type: 'GET'

         });

    });










});
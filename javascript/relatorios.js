$(document).ready(function() {
    if($("#perito_select").length>0) {
        $('#perito_select').select2();
    }
    
    $('#perito_select').on('change', function() {
        let idPerito = $(this).val();
        var formData = new FormData();
        formData.append("idPerito",idPerito);
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/relatorios.php?status=ajax&case=1",
            async: true,
            data: formData,
            cache: false,
            success: function (msg) {

            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data',
            timeout: 60000
        }).done(function (dir){
            openWindowWithPost(location.origin+"/pages/relatorios.php?status=ajax&case=2", {
                dir: dir,
            });
        })
    })
});

function openWindowWithPost(url, data) {
    var form = document.createElement("form");
    form.target = "_blank";
    form.method = "POST";
    form.action = url;
    form.style.display = "none";

    for (var key in data) {
        var input = document.createElement("input");
        input.type = "hidden";
        input.name = key;
        input.value = data[key];
        form.appendChild(input);
    }

    document.body.appendChild(form);
    form.submit();
    document.body.removeChild(form);
}



      // Load the Visualization API and the corechart package.
      google.charts.load('current', {'packages':['corechart']});

      // Set a callback to run when the Google Visualization API is loaded.
      google.charts.setOnLoadCallback(drawChart);

      // Callback that creates and populates a data table,
      // instantiates the pie chart, passes in the data and
      // draws it.
      function drawChart() {


        $.ajax({
            type: "POST",
            url: location.origin+"/pages/relatorios.php?status=ajax&case=1",
            async: true,
            cache: false,
            success: function (msg) {
                msg = JSON.parse(msg);
                if(msg.status == 200){
                    let response = msg.response;
                    debugger
                    var data = new google.visualization.DataTable();
                    data.addColumn('string', 'Topping');
                    data.addColumn('number', 'Slices');
                    data.addRows(response);
                    var options = {'title':msg.title,
                                    'height':400};
    
                    // Instantiate and draw our chart, passing in some options.
                    var chart = new google.visualization.PieChart(document.getElementById('turmas'));
                    chart.draw(data, options);
                }

  
            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data',
            timeout: 60000
        }).done(function (dir){
        })

        $.ajax({
            type: "POST",
            url: location.origin+"/pages/relatorios.php?status=ajax&case=2",
            async: true,
            cache: false,
            success: function (msg) {
                debugger

                msg = JSON.parse(msg);
                if(msg.status == 200){
                    let response = msg.response;
                    var data = new google.visualization.DataTable();
                    data.addColumn('string', 'Topping');
                    data.addColumn('number', 'Slices');
                    data.addRows(response);
                    var options = {'title':msg.title,
                                    'height':400};

                    var chart = new google.visualization.PieChart(document.getElementById('usuarios'));
                    chart.draw(data, options);
                }

  
            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data',
            timeout: 60000
        }).done(function (dir){
        })

        // $.ajax({
        //     type: "POST",
        //     url: location.origin+"/pages/relatorios.php?status=ajax&case=3",
        //     async: true,
        //     cache: false,
        //     success: function (msg) {
        //         debugger

        //         msg = JSON.parse(msg);
        //         if(msg.status == 200){
        //             let response = msg.response;

        //             var data = new google.visualization.DataTable();
        //             data.addColumn('timeofday', 'Time of Day');
        //             data.addColumn('number', 'Motivation Level');
              
        //             data.addRows([
        //               [{v: [8, 0, 0], f: '8 am'}, 1],
        //               [{v: [9, 0, 0], f: '9 am'}, 2],
        //               [{v: [10, 0, 0], f:'10 am'}, 3],
        //               [{v: [11, 0, 0], f: '11 am'}, 4],
        //               [{v: [12, 0, 0], f: '12 pm'}, 5],
        //               [{v: [13, 0, 0], f: '1 pm'}, 6],
        //               [{v: [14, 0, 0], f: '2 pm'}, 7],
        //               [{v: [15, 0, 0], f: '3 pm'}, 8],
        //               [{v: [16, 0, 0], f: '4 pm'}, 9],
        //               [{v: [17, 0, 0], f: '5 pm'}, 10],
        //             ]);
              
        //             var options = {
        //               title: 'Motivation Level Throughout the Day',
        //               hAxis: {
        //                 title: 'Time of Day',
        //                 format: 'h:mm a',
        //                 viewWindow: {
        //                   min: [7, 30, 0],
        //                   max: [17, 30, 0]
        //                 }
        //               },
        //               vAxis: {
        //                 title: 'Rating (scale of 1-10)'
        //               }
        //             };
              
        //             var chart = new google.visualization.ColumnChart(
        //               document.getElementById('chart_div'));
              
        //             chart.draw(data, options);
        //         }

  
        //     },
        //     error: function (error) {
        //         console.log(error);
        //     },
        //     contentType: false,
        //     processData: false,
        //     enctype: 'multipart/form-data',
        //     timeout: 60000
        // }).done(function (dir){
        // })
      }
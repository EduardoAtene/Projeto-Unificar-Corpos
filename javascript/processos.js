if($("#dataTables1").length>0) {
    
    $('#report-adm-consultoria').on('click', function(e){
        Swal.fire({
            html: "Gerando Relatório. Por favor, aguarde...",
            showLoaderOnConfirm: true,
            allowOutsideClick: false,
        })
        Swal.showLoading();
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/relatorios.php?status=ajax&case=6",
            async: true,
            cache: false,
            success: function (data) {  
                var data = JSON.parse(data);
                if(data.status == 200){
                    var $a = $("<a>");
                    $a.attr("href",data.file);
                    $("body").append($a);
                    $a.attr("download",data.fileName+".xlsx");
                    $a[0].click();
                    $a.remove();

                }else{
                    alert("Ocorreu um problema ao emitir o relatório. Por Favor, tende mais tarde!");
                }
                Swal.close();
            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data',
        });
    })

    tableAdm();
}
if($("#dataTables2").length>0) {
    var table = $('#dataTables2').DataTable({
        "ajax": {
            "url": $('#dataTables2').attr("ajaxpath"),
            "type": "POST"
        },
        columnDefs: [
            { targets: [5],     data: 5, render: {_: 'display',sort: 'order'}},
            { targets: [6],     data: 6, render: {_: 'display',sort: 'order'}},
            {"className": "dt-center", "targets": "_all"}
        ],
        "pageLength":10,
        "serverSide": true,
        "responsive": true,
        "processing": true,
        
    });
}
if($("#dataTablesPerito").length>0) {
    tablePerito("create");
}
if($("#dataTables3").length>0) {
    tableAnalista();
}
$(document).ready(function() {

    if($("#filtraVara_select").length>0) {
        $('#filtraVara_select').select2();
    }
    if($("#filtraTemperatura_select").length>0) {
        $('#filtraTemperatura_select').select2();
    }
    if($("#filtraTemperaturaManual_select").length>0) {
    $('#filtraTemperaturaManual_select').select2();
    }

    $(document).on("change",".tempManual",function(e) {
        let nameProcesso = $(this).closest('tr').find('a[input-processo]')[0].attributes['input-processo'].value;
        var hash = ($(this).closest('tr').find('a[input-hash]')[0].attributes['input-hash'].value).split(".");
        let tempManual = $('#tempManual'+hash[0]+'_select').val();
        let name = $('#tempManual'+hash[0]+'_select option:selected').text().trim();


        if(tempManual == 1){
            var title = "Alteração para Pago";
            var html = "<span style='text-align: center'>Estarás mudando o status do processo <br><b>"+nameProcesso+"</b><br>para <b>"+name+"</b></span><br>"+
                       "<div class='form-group'><label class='form-group d-flex p-2 bd-highlight' style='margin-bottom:0;margin-top:1rem' for='valuePay'>Informe o Valor Pago</label>"+//<br><input id='swal-input1' class='form-control'><br> "+
                        "<div class='input-group mb-3'><div class='input-group-prepend'>"+
                           "<span class='input-group-text' style='color: #121212;'>R$</span>"+
                        "</div><input type='text' class='form-control' placeholder='Valor' aria-label='Valor' aria-describedby='basic-addon1' id='valuePay'></div></div>"+
                        "<div class='form-group'><label class='form-group d-flex p-2 bd-highlight' style='margin-bottom:0;margin-top:1rem' >Informe a Data do Pagamento</label><br>"+//<br><input id='swal-input2' class='form-control'>"+
                        "<input id='date' class='form-control' type='date'></div>Continuar?";
        }else{
            var title = "Alteração de Temperatura";
            var html = "<span style='text-align: center'>Estarás mudando o status do processo <br><b>"+nameProcesso+"</b><br>para <b>"+name+"</b></span><br> Continuar?";
        }

        Swal.fire({
            title: title,
            html: html,
            icon: 'info',
            showCancelButton: true,
            confirmButtonText: 'Alterar',
            confirmButtonColor: "#00FF00",
            allowOutsideClick: true,
            preConfirm: function() {
             return new Promise(function(resolve,reject) {
                if(tempManual == 1){
                    var valuePay = $("#valuePay").val();
                    var valueDate = $("#date").val().length;
                    var pode = true;
                    if(valuePay == ''){
                        Swal.showValidationMessage("Por favor, informe o Valor Pago");
                        pode = false;
                        if(valueDate <10){
                            Swal.showValidationMessage("Por favor, informe o Periodo de Pagamento");
                            pode = false;
                        }
                    }
                    if(pode)
                        resolve(true);

                }else
                    resolve(true);
             })
            },
            didOpen: function(valuePay) {
                $('#valuePay').keyup(function(){
                    formatarMoeda('valuePay');
                });

            }
        }).then((result) => {    
            if(result['isConfirmed']){

                var formData = new FormData();
                formData.append("tempManual", tempManual);
                formData.append("hash",hash[0]);
                if(tempManual == 1){
                    formData.append("value",$("#valuePay").val());
                    formData.append("periodo",$("#date").val());
                }
                $.ajax({
                    type: "POST",
                    url: location.origin+"/pages/processos.php?status=ajax&case=15",
                    async: true,
                    data: formData,
                    cache: false,
                    success: function (data) {
                        debugger
                        data = data.split("|");
                        if(data[0] == 1){
                            alert ("Temperatua Manual Alterada!");
                            if(data[1] == 1){
                                $('#tempManual'+hash[0]).remove();
                                $('#tempManual'+hash[0]+"_select").remove();
                                $("input[input-hash="+hash[0]+"]").remove();
                                $('#div_'+hash[0]).html("PAGO");
                                $('#pay_'+hash[0]).html(data[2]);
                                $('#check_'+hash[0]).html("-");

                                
                                //$('#tempManual'+hash+"_select").prop('disabled', 'disabled');

                            }
                            else
                                $('#tempManual'+hash).val(tempManual);

                        }
                        //selectTablePerito(listCheckValues);
                    },
                    error: function (error) {
                        console.log(error);
                    },
                    contentType: false,
                    processData: false,
                    enctype: 'multipart/form-data',
                });
            }
            else{
                var valueOld = $('#tempManual'+hash[0]).val();
                $('#tempManual'+hash[0]+'_select').val(valueOld);
            }

        });    

    })

    $('#filtraVara_select').on('change', function(e) {
        tablePerito('reload');
    })
    $('#filtraTemperatura_select').on('change', function(e) {
        tablePerito('reload');
    })
    $('#filtraTemperaturaManual_select').on('change', function(e) {
        tablePerito('reload');
    })
    $('#report-perito').on('click', function(e){
        Swal.fire({
            html: "Gerando Relatório. Por favor, aguarde...",
            showLoaderOnConfirm: true,
            allowOutsideClick: false,
        })
        Swal.showLoading();
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/relatorios.php?status=ajax&case=4",
            async: true,
            cache: false,
            success: function (data) { 
                var data = JSON.parse(data);
                if(data.status == 200){
                    var $a = $("<a>");
                    $a.attr("href",data.file);
                    $("body").append($a);
                    $a.attr("download",data.fileName+".xlsx");
                    $a[0].click();
                    $a.remove();

                }else{
                    alert("Ocorreu um problema ao emitir o relatório. Por Favor, tende mais tarde!");
                }
                Swal.close();
            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data',
        });
    })
    $(this).on("click", "#checkConsultoria", function(){
        var hash = $(this).attr("input-hash");
        var formData = new FormData();
        formData.append("hash",hash);
        formData.append("listHashChecked",$("#valuesCheckElement").val());
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/processos.php?status=ajax&case=14",
            async: false,
            data: formData,
            cache: false,
            success: function (data) {
                data = JSON.parse(data);       

                $("#valuesCheckElement").val(data['data']);
                if(data['count']>0){
                    $("#enviarConsultoria")[0].style.display = "";
                    $("#alert-processos").text(data['count']+' Processos Selecionados = R$ '+data['value']);
                    $("#valorUnitario").text(data['valueUnique']);//+data['value']);
                    if(data['desconto'] == true){
                        $("#descontoAplicado").html("<span class='mdi mdi-information-outline blue' style='color:blue' title='Desconto Aplicado por conta que foi enviado para consultoria "+ data['allProcessSend'] +" processo esse mês'></span>");
                    }else{
                        $("#descontoAplicado").html(""); 
                    }
                }else{
                    $("#enviarConsultoria")[0].style.display = "none";
                    $("#alert-processos").text('');
                } 
            },
            contentType: false,
            processData: false,
        });

    })
    $('#enviarConsultoria').on('click', function(e){
        var formData = new FormData();
        formData.append("hash",-1);
        formData.append("listHashChecked",$("#valuesCheckElement").val());
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/processos.php?status=ajax&case=14",
            async: false,
            data: formData,
            cache: false,
            success: function (data) {
                data = JSON.parse(data);   

                var text = "Foram selecionados "+data['count']+" processos.<br>São <b>"+data['valueUnique']+"</b> reais por processo. <br><br>Isso resultará em um acréscimo do valor base da próxima mensalidade em <b>"+data['value']+"</b> reais.<br><br>Continuar?";
                Swal.fire({
                    title: 'Solicitação de Consultoria',
                    html: text,
                    icon: 'question',
                    showCancelButton: true,
                    confirmButtonText: 'SIM',
                    confirmButtonColor: "#00FF00"
                }).then((result) => {
                    if(result['isConfirmed'] == true){
                        /*destruirTable("dataTablesPerito")
                        var hash = [];
                        $.each($("input[name='consultoria']:checked"), function(){
                            hash.push($(this).attr("input-hash"));
                        });
                        tablePerito();*/
        
                        var formData = new FormData();
                        formData.append("listHashChecked",$("#valuesCheckElement").val());

                        $.ajax({
                            type: "POST",
                            url: location.origin+"/pages/processos.php?status=ajax&case=16",
                            data: formData, 
                            processData: false,
                            contentType: false,
                            success: function (data) {   
                                alert ("Enviado a Consultoria!");
                                setTimeout(function () { location.reload(); }, 2000);
                            }
                        })
                    }
                })
            },
            contentType: false,
            processData: false,
        })

    })

    $(this).on("click", ".modalProcesso", function(){
        
        var hash = $(this).attr("input-hash");

        var formData = new FormData();
        formData.append("processo",hash);
        $.ajax({
            type: "POST",
            url: location.origin+"/pages/processos.php?status=ajax&case=12",
            async: true,
            data: formData,
            cache: false,
            success: function (msg) {
                
                msg = JSON.parse(msg);
                if(msg.status == 200){
                    let modalProcesso = $('#modalProcesso');
                    modalProcesso.find('.modal-header').html(msg.header);
                    modalProcesso.find('.modal-body').html(msg.body);
                    modalProcesso.find('.modal-footer').html(msg.footer);
                    modalProcesso.modal('show');

                }else{
                    Swal.fire({
                        title: 'Erro!',
                        text: msg.message,
                        icon: 'error',
                        confirmButtonText: 'Ok'
                    })
                }
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data'
        });
    })

})
function modaledicao(hash,type){
    
    var formData = new FormData();
    formData.append("processo", hash);
    let modalProcesso = $('#modalProcesso');

    $.ajax({
        type: "POST",
        url: location.origin+"/pages/processos.php?status=ajax&case=12",
        async: true,
        data: formData,
        cache: false,
        success: function (msg) {         
            msg = JSON.parse(msg);
            if(msg.status == 200){
                
                modalProcesso.find('.modal-header').html(msg.header);
                modalProcesso.find('.modal-body').html(msg.body);
                modalProcesso.find('.modal-footer').html(msg.footer);
                modalProcesso.modal('show');
                if(msg.inputElements){
                    $('#tempAnalise_select').select2();
                    /*var cleave = new Cleave('#valueAnalise', {
                        delimiters: ['.','.','.', '.', ','],
                        blocks: [3, 3, 3, 2],
                        uppercase: true
                    });*/
                    $('#valueAnalise').keyup(function(){
                        formatarMoeda('valueAnalise');
                    });
                
                }
                $("button[name = historico-consultoria]").on("click",function (e) {
                    
                    var hash = $(this).attr("input-hash");
                    //var processo = $(this).attr("input-processo");

                    var formData = new FormData();
                    formData.append("processo",hash);
                    $.ajax({
                        type: "POST",
                        url: location.origin+"/pages/processos.php?status=ajax&case=19",
                        async: true,
                        data: formData,
                        cache: false,
                        success: function (msg) {
                            msg = JSON.parse(msg);
                            if(msg.status == 200){
                                /* Modal Consultoria e não swal
                                let modalProcesso = $('#modalHistoricoConsultoria');
                                modalProcesso.find('.modal-header').html(msg.header);
                                modalProcesso.find('.modal-body').html(msg.body);
                                modalProcesso.find('.modal-footer').html(msg.footer);
                                modalProcesso.modal('show');
                                */
                                Swal.fire({
                                    html: msg.body,
                                    icon: 'info'
                                })
                            }else{
                                Swal.fire({
                                    title: 'Erro!',
                                    text: msg.message,
                                    icon: 'error',
                                    confirmButtonText: 'Ok'
                                })
                            }
                        },
                        contentType: false,
                        processData: false,
                        enctype: 'multipart/form-data'
                    });
                })

                $("a[name = alterar-info-processo]").on("click",function (e) {
                    var type = $(this).attr("data-type");
                    var html = "<textarea id='text' class='form-control' type='text' ></textarea>";
                    if(type == 'autor'){
                        title = "Alterar autor do Processo";
                    }else if(type == 'reu'){
                        title = "Alterar réu do Processo";
                    }

                    Swal.fire({
                        title: title,
                        html: html,
                        icon: 'info',
                        showCancelButton: true,
                        confirmButtonText: 'Alterar',
                        confirmButtonColor: "#00FF00",
                        allowOutsideClick: true,
                        preConfirm: function() {
                         return new Promise(function(resolve,reject) {
                            
                            var text = $("#text").val();
                            var pode = true;
                            if(text == ''){
                                Swal.showValidationMessage("Por favor, informe o Valor Pago");
                                pode = false;
                            }
                            if(pode)
                                resolve(true);
                         })
                        },
                        didOpen: function(valuePay) {
                            if(title.includes('autor')){
                                text = $('a[data-type = autor]').attr("data-textA");
                            }if(title.includes('réu')){
                                text =  $('a[data-type = reu]').attr("data-textR");
                            }  
                            $('#text').val(text.toUpperCase());

                            $('#text').keyup(function(){
                                $('#text').val($("#text").val().toUpperCase());
                            })
    
                        }
                    }).then((result) => {    
                        
                        if(result['isConfirmed']){
                            var text = $('#text').val();
                            var hashP = $("a[name = alterar-info-processo]").attr("data-hash");
                            var formData = new FormData();
                            formData.append("hash", hashP);
                            formData.append("text", text);
                            formData.append("type",type);

                            $.ajax({
                                type: "POST",
                                url: location.origin+"/pages/processos.php?status=ajax&case=20",
                                async: true,
                                data: formData,
                                cache: false,
                                success: function (data) {
                                    Swal.fire({
                                        title: 'Sucesso!',
                                        text: "Alterado com sucesso!",
                                        icon: 'success',
                                        confirmButtonText: 'Ok'
                                    }).then((result) => {
                                        if(type == 'autor'){
                                            $("#eAutor").text(text);
                                            $('a[data-type = autor]').attr("data-textR",text);
                                        }else if(type == 'reu'){
                                            $("#eReu").text(text);
                                            $('a[data-type = reu]').attr("data-textR",text);
                                        }
                                        
                                    })
                                },
                                error: function (error) {
                                    console.log(error);
                                },
                                contentType: false,
                                processData: false,
                                enctype: 'multipart/form-data',
                            });
                        }
                        else{
                            var valueOld = $('#tempManual'+hash).val();
                            $('#tempManual'+hash+'_select').val(valueOld);
                        }
            
                    });  
                })
                
                
                $("a[name = change-file]").on("click",function (e) {
                    var type = $(this).attr("type");
                    var text = "";
                    if(type=="arquivoReferencia")
                        text = "Arquivo Referência";
                    else if(type == "acordo")
                        text = "Acordo";
                    else if(type == "patrimonial")
                        text = "Patrimonial";
                    else if(type == "despacho")
                        text = "Último Despacho";

                    $("#arquivoHidden").prop("tipoArquivo",type);

                    Swal.fire({
                        title: "Alteração no "+text,
                        icon: 'info',
                        showCancelButton: true,
                        showDenyButton: true,
                        confirmButtonText: 'Upload Arquivo',
                        denyButtonText: 'Excluir Arquivo',
                        confirmButtonColor: "#00FF00",
                        allowOutsideClick: true,
                        backdrop: true
                    }).then((result) => { 
                        if(result['isConfirmed']){                            
                            $("#arquivoHidden").click();
                        }
                        else if(result['isDenied']){
                            var hashExtrac = $("#arquivoHidden").attr("hash-extrac");
                            var tipoArquivo = $("#arquivoHidden").prop("tipoArquivo");
                            formData.append("type", "changeFileExtracao");
                            formData.append("tipoArquivo", tipoArquivo);
                            formData.append("tipoAcao", "excluir");
                            formData.append("hash", hashExtrac);  
                            $.ajax({
                                type: "POST",
                                url: location.origin+"/pages/processos.php?status=ajax&case=7",
                                async: true,
                                data: formData,
                                cache: false,
                                success: function (data) {                     
                                    data = JSON.parse(data);
                                    if(data['status']=='200'){
                                        Swal.fire({
                                            title: 'Arquivo Removido!',
                                            icon: 'success',
                                            confirmButtonText: 'Ok'
                                        })    
                                        $("#"+tipoArquivo+"-file").removeAttr('href');
                                        $("#"+tipoArquivo+"-file").removeAttr('target');
                                    }
                                    else{
                                        Swal.fire({
                                            title: 'Ocorreu um erro. Por favor, tente mais tarde!',
                                            text: data.message,
                                            icon: 'error',
                                            confirmButtonText: 'Ok'
                                        })     
                                    }                
                                },
                                error: function (error) {
                                    console.log(error);
                                },
                                contentType: false,
                                processData: false,
                                enctype: 'multipart/form-data',
                            });
                        }
            
                    });  


                    
                })
                
                $("#arquivoHidden").on("change",function (e) {
                    var fileExtracao = $("#arquivoHidden")[0].files;
                    var hashExtrac = $(this).attr("hash-extrac");
                    var tipoArquivo = $(this).prop("tipoArquivo");

                    if(fileExtracao.length>0){
                        var formData = new FormData();
                        formData.append('files', fileExtracao[0]);	
                        formData.append("type", "changeFileExtracao");
                        formData.append("tipoArquivo", tipoArquivo);
                        formData.append("tipoAcao", "anexar");
                        formData.append("hash", hashExtrac);
                                                    
                        $.ajax({
                            type: "POST",
                            url: location.origin+"/pages/processos.php?status=ajax&case=7",
                            async: true,
                            data: formData,
                            cache: false,
                            success: function (data) {
                                data = JSON.parse(data);
                                if(data['status']=='200'){
                                    Swal.fire({
                                        title: 'Arquivo Inserido!',
                                        text: data.message,
                                        icon: 'success',
                                        confirmButtonText: 'Ok'
                                    })    
                                    $("#"+tipoArquivo+"-file").attr("href", data.linkRef);
                                    $("#"+tipoArquivo+"-file").attr("target",'_blank');
                                }
                                else{
                                    Swal.fire({
                                        title: 'Ocorreu um erro. Por favor, tente mais tarde!',
                                        text: data.message,
                                        icon: 'error',
                                        confirmButtonText: 'Ok'
                                    })     
                                }                
                            },
                            error: function (error) {
                                console.log(error);
                            },
                            contentType: false,
                            processData: false,
                            enctype: 'multipart/form-data',
                        });
                    }else{
                        Swal.fire({
                            title: 'Arquivo não anexado!',
                            icon: 'error',
                            confirmButtonText: 'Ok'
                        })     
                    }
                })

                if(type == "adm"){
                    $(document).on("change", "#peticao-comprov", function(){
                        var filesAdm = $("#peticao-comprov")[0].files;
                        var hashConsultoriaAdm = $(this).attr("data-hash-consultoria");

                        if(filesAdm.length>0){
                            var formData = new FormData();
                            formData.append('files', filesAdm[0]);	
                            formData.append("type", "changeFileAnalista");
                            formData.append("hashConsultoria", hashConsultoriaAdm);
                                                        
                            $.ajax({
                                type: "POST",
                                url: location.origin+"/pages/processos.php?status=ajax&case=7",
                                async: true,
                                data: formData,
                                cache: false,
                                success: function (data) {
                                    
                                    data = JSON.parse(data);
                                    if(data['status']=='200'){
                                        $("#fileNamePeticao").text(data['fileNameComprimido']);
                                        $('#linkPeticao')[0].attributes.title.value = data['fileName'];
                                        $("#linkPeticao")[0].style.display = "";
                                        $("#lixeiraPeticao")[0].style.display = "";
                                        $("#peticao-comprov").val('');
                                    }
                                    else{
                                        Swal.fire({
                                            title: 'Ocorreu um erro. Por favor, tente mais tarde!',
                                            text: data.message,
                                            icon: 'error',
                                            confirmButtonText: 'Ok'
                                        })     
                                    }                
                                },
                                error: function (error) {
                                    console.log(error);
                                },
                                contentType: false,
                                processData: false,
                                enctype: 'multipart/form-data',
                            });
                        }
                        
                    }); 	
                }
            }
            else{
                Swal.fire({
                    title: 'Erro!',
                    text: msg.message,
                    icon: 'error',
                    confirmButtonText: 'Ok'
                })
            }
        },
        error: function (error) {
            console.log(error);
        },
        contentType: false,
        processData: false,
        enctype: 'multipart/form-data'
    }).done(function (data) {
        $("button[data-typeform = edit-processo]").on("click",function (e) {
            let type = $(this).attr('data-type');
            
            switch (type) {
                case 'assumir':
                    let hashConsultoria = $(this).attr('data-hash-consultoria');

                    var formData = new FormData();
                    formData.append("type", type);
                    formData.append("hashConsultoria", hashConsultoria);
                    $.ajax({
                        type: "POST",
                        url: location.origin+"/pages/processos.php?status=ajax&case=7",
                        async: true,
                        data: formData,
                        cache: false,
                        success: function (data) {
                            
                            data = JSON.parse(data);
                            if(data['status']=='200'){
                                Swal.fire({
                                    title: 'Sucesso!',
                                    text: data.message,
                                    icon: 'success',
                                    confirmButtonText: 'Ok'
                                }).then((result) => {
                                    console.log(result);
                                    if (result.isConfirmed) {
                                        modaledicao(hash,"analista");
                                    }
                                })
                            }
                            else{
                                console.log(data['message']);
                                alert('Ocorreu um erro durante a importação do arquivo');
                            }                
                        },
                        error: function (error) {
                            console.log(error);
                        },
                        contentType: false,
                        processData: false,
                        enctype: 'multipart/form-data',
                    });
                break;
                case 'saveValidacao':
                    
                    //let idConsultoria = $(this).attr('data-idConsultoria');
                    var hashConsultoriaa = $(this).attr('data-hash-consultoria');

                    let novaTemperatura = $("#tempAnalise_select").val();
                    let justificativa = $("#justificativa").val();
                    let idAnalise = $("#idAnalise").val();
                    let valueAnalise = $("#valueAnalise").val();
                    var files = $("#peticao-comprov")[0].files;
                    var pode = true;

                    if(novaTemperatura == 0 || novaTemperatura == "" ){
                        $("#tempAnalise").removeClass("is-valid").addClass("is-invalid"); // No Working
                        pode = false;
                    }else
                        $("#tempAnalise").removeClass("is-invalid").addClass("is-valid");

                    if(justificativa == null || justificativa == "" ){
                        $("#justificativa").removeClass("is-valid").addClass("is-invalid");
                        pode = false;
                    }else
                        $("#justificativa").removeClass("is-invalid").addClass("is-valid");

                    if( ($("#valueAnalise").prop("required") && (valueAnalise == null || valueAnalise == "" ))){
                        $("#valueAnalise").removeClass("is-valid").addClass("is-invalid");
                        pode = false;
                    }else
                        $("#valueAnalise").removeClass("is-invalid").addClass("is-valid");

                    $("#idAnalise").removeClass("is-invalid").addClass("is-valid");
                    $("#files").removeClass("is-invalid").addClass("is-valid");


                    if(pode){
                        var formData = new FormData();
                        formData.append("type", type);
                        formData.append("hashConsultoria", hashConsultoriaa);
                        formData.append("temperatura", novaTemperatura);
                        formData.append("justificativa", justificativa);
                        formData.append("id", idAnalise);
                        formData.append("valor", valueAnalise);

                        if(files.length>0){
                            formData.append('files', files[0]);	
                        }
                        
                        $.ajax({
                            type: "POST",
                            url: location.origin+"/pages/processos.php?status=ajax&case=7",
                            async: true,
                            data: formData,
                            cache: false,
                            success: function (data) {
                                
                                data = JSON.parse(data);
                                if(data['status']=='200'){
                                    Swal.fire({
                                        title: 'Sucesso!',
                                        text: data.message,
                                        icon: 'success',
                                        confirmButtonText: 'Fechar'
                                    }).then((result) => {
                                         modaledicao(hash,"analista");
                                    })
                                }
                                else{
                                    console.log(data['message']);
                                    alert('Ocorreu um erro durante a importação do arquivo');
                                }                
                            },
                            error: function (error) {
                                console.log(error);
                            },
                            contentType: false,
                            processData: false,
                            enctype: 'multipart/form-data',
                        });
                    }//else
                        //alert('Por favor, preencha os campos obrigatórios'); // Alerta fica atrás do modal '-' Verificar Marcos

                break;
                case 'saveValidacaoAdm':

                    let hashConsultoriaAdm = $(this).attr('data-hash-consultoria');
                    let novaTemperaturaAdm = $("#tempAnalise_select").val();
                    let justificativaAdm = $("#justificativa").val();
                    let idAnaliseAdm = $("#idAnalise").val();
                    let valueAnaliseAdm = $("#valueAnalise").val();
                    var pode = true;

                    if(novaTemperaturaAdm == 0 || novaTemperaturaAdm == "" ){
                        $("#tempAnalise").removeClass("is-valid").addClass("is-invalid"); // No Working
                        pode = false;
                    }else
                        $("#tempAnalise").removeClass("is-invalid").addClass("is-valid");

                    if(justificativaAdm == null || justificativaAdm == "" ){
                        $("#justificativa").removeClass("is-valid").addClass("is-invalid");
                        pode = false;
                    }else
                        $("#justificativa").removeClass("is-invalid").addClass("is-valid");

                    if( ($("#valueAnalise").prop("required") && (valueAnaliseAdm == null || valueAnaliseAdm == "" ))){
                        $("#valueAnalise").removeClass("is-valid").addClass("is-invalid");
                        pode = false;
                    }else
                        $("#valueAnalise").removeClass("is-invalid").addClass("is-valid");

                    $("#idAnalise").removeClass("is-invalid").addClass("is-valid");
                    $("#files").removeClass("is-invalid").addClass("is-valid");


                    if(pode){
                        var formData = new FormData();
                        formData.append("type", type);
                        formData.append("hashConsultoria", hashConsultoriaAdm);
                        formData.append("temperatura", novaTemperaturaAdm);
                        formData.append("justificativa", justificativaAdm);
                        formData.append("id", idAnaliseAdm);
                        formData.append("valor", valueAnaliseAdm);

                        $.ajax({
                            type: "POST",
                            url: location.origin+"/pages/processos.php?status=ajax&case=7",
                            async: true,
                            data: formData,
                            cache: false,
                            success: function (data) {
                                
                                data = JSON.parse(data);
                                if(data['status']=='200'){
                                    Swal.fire({
                                        title: 'Sucesso!',
                                        text: data.message,
                                        icon: 'success',
                                        confirmButtonText: 'Fechar'
                                    }).then((result) => {
                                        modaledicao(hash,"adm");
                                        
                                    })
                                }
                                else{
                                    console.log(data['message']);
                                    alert('Ocorreu um erro durante a importação do arquivo');
                                }                
                            },
                            error: function (error) {
                                console.log(error);
                            },
                            contentType: false,
                            processData: false,
                            enctype: 'multipart/form-data',
                        });
                    }//else
                     //   alert('Por favor, preencha os campos obrigatórios'); // Alerta fica atrás do modal '-' Verificar Marcos

                break;
                case 'salvar':
                    let vencimento = $("#data-vencimento").val();
                    let valorArbitrado = $("#valor-arbitrado").val();
                    let valorPago = $("#valor-pago").val();
                    let comentarios = $("#comentarios").val();
                    if($("#select-ordem_select").length>0)
                        ordem = $("#select-ordem_select").val();
                    if($("#select-expresso_select").length>0)
                        expresso = $("#select-expresso_select").val();
                    let processo = $("#status-processo_select").val();
                    let peticao = $("#peticao_select").val();
                    let peticao_required = $("#peticao").prop('required');
                    tudoOk = true;
                    if($("#select-ordem_select").length>0){
                        if(ordem=='' || ordem==null){
                            tudoOk = false;
                            let aux = $("#select-ordem_select").parent().find('.select2-selection--single');
                            aux.attr('style','border: 1px solid red;');
                        }
                        else{
                            let aux = $("#select-ordem_select").parent().find('.select2-selection--single');
                            aux.attr('style','border: 1px solid green;');
                        }
                    }
                    if($("#select-expresso_select").length>0){
                        if(expresso=='' || expresso==null){
                            tudoOk = false;
                            let aux = $("#select-expresso_select").parent().find('.select2-selection--single');
                            aux.attr('style','border: 1px solid red;');
                        }
                        else{
                            let aux = $("#select-expresso_select").parent().find('.select2-selection--single');
                            aux.attr('style','border: 1px solid green;');
                            if(expresso==1){
                                if(valorPago==''||valorPago==null){
                                    if(valorPago == ''){
                                        tudoOk = false;
                                        $("#valor-pago").attr('class','form-control is-invalid')
                                    }
                                    else{
                                        $("#valor-pago").attr('class','form-control is-valid')
                                    }
                                }
                            }
                        }
                    }
                    if(processo=='' || processo==null){
                        tudoOk = false;
                        let aux = $("#status-processo").parent().find('.select2-selection--single');
                        aux.attr('style','border: 1px solid red;');
                    }
                    else{
                        let aux = $("#status-processo").parent().find('.select2-selection--single');
                        aux.attr('style','border: 1px solid green;');
                    }
                    if(peticao=='' && peticao_required){
                        tudoOk = false;
                        let aux = $("#peticao_select").parent().find('.select2-selection--single');
                        aux.attr('style','border: 1px solid red;');
                    }
                    else{
                        let aux = $("#peticao_select").parent().find('.select2-selection--single');
                        aux.attr('style','border: 1px solid green;');
                    }
                    if(tudoOk){
                        var formData = new FormData();
                        formData.append("type", type);
                        formData.append("hash", hash);
                        formData.append("valorArbitrado", valorArbitrado);
                        if(expresso==1)
                            formData.append("valorPago", valorPago);
                        formData.append("comentarios", comentarios);
                        formData.append("vencimento", vencimento);
                        if($("#select-ordem_select").length>0)
                            formData.append("ordem", ordem);
                        if($("#select-expresso_select").length>0)
                            formData.append("expresso", expresso);
                        formData.append("processo", processo);
                        if(peticao_required)
                            formData.append("peticao", peticao);
                        $.ajax({
                            type: "POST",
                            url: location.origin+"/pages/processos.php?status=ajax&case=7",
                            async: true,
                            data: formData,
                            cache: false,
                            success: function (data) {
                                data = JSON.parse(data);
                                if(data['status']=='200'){
                                    Swal.fire({
                                        title: 'Sucesso!',
                                        text: data.message,
                                        icon: 'success',
                                        confirmButtonText: 'Ok'
                                    }).then((result) => {
                                        console.log(result);
                                        if (result.isConfirmed) {
                                            modaledicao(hash,"inconclusiv");
                                        }
                                    })
                                }
                                else{
                                    console.log(data['message']);
                                    alert('Ocorreu um erro durante a importação do arquivo');
                                }                
                            },
                            error: function (error) {
                                console.log(error);
                            },
                            contentType: false,
                            processData: false,
                            enctype: 'multipart/form-data',
                        });
                    }
                    else{
                        alert('Favor preencher todos os campos corretamente');
                    }

                break;
            }
        })

        $('#tempAnalise_select').on("change", function(){
            $tempAnalise = $('#tempAnalise_select').val();
            if($tempAnalise == 1){
                $("#valueAnalise").prop("required",true);
            }else{
                $("#valueAnalise").prop("required",false);
            }

        });
        
        $("a[name = removePeticao]").on("click",function () {
            Swal.fire({
                title: 'Tem certeza que deseja remover?',
                icon: 'info',
                showCancelButton: true,
                confirmButtonText: 'Remover',
                confirmButtonColor: "#00FF00",
            }).then((result) => {
                if (result.isConfirmed) {   
                    var hash = $(this).attr('input-hash');
                    var type = $(this).attr('data-type');
                    var formData = new FormData();
                    formData.append("hash",hash);
                    formData.append("type", type);               
                    $.ajax({
                        type: "POST",
                        url: location.origin+"/pages/processos.php?status=ajax&case=7",
                        async: true,
                        cache: false,
                        data: formData,
                        success: function (msg) { 
                            msg = JSON.parse(msg);
                            if(msg.status == 200){
                                $("#fileNamePeticao").text("");
                                $('#linkPeticao')[0].attributes.title.value = " ";
                                $("#linkPeticao")[0].style.display = "none";
                                $("#lixeiraPeticao")[0].style.display = "none";
                            }else{
                                Swal.fire({
                                    title: 'Erro!',
                                    text: msg.message,
                                    icon: 'error',
                                    confirmButtonText: 'Ok'
                                })
                            }
                        },
                        error: function (error) {
                            Swal.fire({
                                title: 'Erro!',
                                text: "Não foi possível remover a petição. Por favor, tente mais tarde!",
                                icon: 'error',
                                confirmButtonText: 'Ok'
                            })
                        },
                        contentType: false,
                        processData: false,
                        enctype: 'multipart/form-data'
                    });
                }
            })

        })

        $('#closeModalA').click(function() {
            if(type == "adm")
                tableAdm(false);
            else
                tableAnalista(false);
        });

    });
}


var Upload = function (file) {
    this.file = file;
};

Upload.prototype.getType = function() {
    return this.file.type;
};
Upload.prototype.getSize = function() {
    return this.file.size;
};
Upload.prototype.getName = function() {
    return this.file.name;
};
Upload.prototype.doUpload = function () {
    var that = this;
    var formData = new FormData();
    var hash = $("button[data-typeForm=edit-processo]").attr('data-hash');

    // add assoc key values, this will be posts values
    //formData.append("file", this.file, this.getName());
    formData.append("files", $('#file')[0].files[0]);
    formData.append("upload_file", true);
    formData.append("hash", hash);

    $.ajax({
        type: "POST",
        url: location.origin+"/pages/processos.php?status=ajax&case=5",
        xhr: function () {
            var myXhr = $.ajaxSettings.xhr();
            if (myXhr.upload) {
                myXhr.upload.addEventListener('progress', that.progressHandling, false);
            }
            return myXhr;
        },
        success: function (data) {
            data = JSON.parse(data);
            if(data['status']=='200'){
                console.log('moranguinho');
                Swal.fire({
                    title: 'Sucesso!',
                    text: data.message,
                    icon: 'success',
                    confirmButtonText: 'Ok'
                })
            }
            else{
                console.log(data['message']);
                alert('Ocorreu um erro durante a importação do arquivo');
            }
        },
        error: function (error) {
            console.log(error);
        },
        async: true,
        data: formData,
        cache: false,
        contentType: false,
        processData: false,
        enctype: 'multipart/form-data',
        timeout: 60000
    });
    
};

Upload.prototype.progressHandling = function (event) {
    // Assuming we have an empty <div id="container"></div> in
    // HTML
    var bar = new ProgressBar.Line('#container', {easing: 'easeInOut'});
    $("#container").show();
    var percent = 0;
    var position = event.loaded || event.position;
    var total = event.total;
    // var progress_bar_id = "#progress-wrp";
    if (event.lengthComputable) {
        percent = Math.ceil(position / total * 100);
    }

    bar.animate(percent/100);  // Value from 0.0 to 1.0
    // update progressbars classes so it fits your code
    // $(progress_bar_id + " .progress-bar").css("width", +percent + "%");
    // $(progress_bar_id + " .status").text(percent + "%");
};


function updateSelect(peticao,processo,hashProcesso){
    var formData = new FormData();
    formData.append("processo", processo);
    formData.append("hashProcesso", hashProcesso);
    var retorno
    $.ajax({
        type: "POST",
        url: location.origin+"/pages/processos.php?status=ajax&case=8",
        async: true,
        data: formData,
        cache: false,
        success: function (data) {
            data = JSON.parse(data);
            if(data['status']=='200'){
                peticao.select2('destroy').empty().select2({data: data['message']});
                peticao.siblings(".select2-container").show()
                peticao.siblings(".control-label").show()
                peticao.siblings("input").prop('required',true);
            }
            else{
                peticao.siblings(".select2-container").hide()
                peticao.siblings(".control-label").hide()
                peticao.siblings("input").prop('required',false);
            }                
        },
        error: function (error) {
            console.log(error);
        },
        contentType: false,
        processData: false,
        enctype: 'multipart/form-data',
    });
}

function updateSelect2(typeFilter,filter,arrayToFilter,hashProcesso){
    var formData = new FormData();
    formData.append("typeFilter", typeFilter);
    formData.append("filter", JSON.stringify(filter));
    formData.append("arrayToFilter", JSON.stringify(arrayToFilter["filtros"]));
    formData.append("hashProcesso", hashProcesso);
    var retorno
    $.ajax({
        type: "POST",
        url: location.origin+"/pages/processos.php?status=ajax&case=10",
        async: true,
        data: formData,
        cache: false,
        success: function (data) {
            data = JSON.parse(data);
            let elementTemp;
            Object.keys(data.message).forEach(elementName => {
                elementTemp = arrayToFilter["selects"][elementName];
                elementTemp.select2('destroy').empty().select2({data: data['message'][elementName]}).val([]).trigger('change',['code']);
                if(data['message'][elementName].length>0){
                    elementTemp.siblings(".select2-container").show()
                    elementTemp.siblings(".control-label").show()
                    elementTemp.siblings("input").prop('required',true);
                }
                else{
                    elementTemp.siblings(".select2-container").hide()
                    elementTemp.siblings(".control-label").hide()
                    elementTemp.siblings("input").prop('required',false);
                }

            });
            if(typeFilter=="expresso"){
                if(filter.expresso==1){
                    $("#valor-pago").show();
                    $("#valor-pago").siblings("label").show();
                }
                else{
                    $("#valor-pago").hide();
                    $("#valor-pago").siblings("label").hide();
                }
            }       
        },
        error: function (error) {
            console.log(error);
        },
        contentType: false,
        processData: false,
        enctype: 'multipart/form-data',
    });

}

function destruirTable(tableName){
    var tabela = $('#'+tableName+'').DataTable();
    tabela.destroy();
}

function tablePerito(type){
    if(type != 'create'){
        var table = $("#dataTablesPerito").DataTable();
        table.clear();
        table.ajax.reload();
    }else{
        var table = $('#dataTablesPerito').DataTable({
            "ajax": {
                "url": $('#dataTablesPerito').attr("ajaxpath"),
                "type": "POST",
                "data": function(d) {
                    d.vara = $('#filtraVara_select').val();
                    d.temperatura = $('#filtraTemperatura_select').val();
                    d.temperaturaManual = $('#filtraTemperaturaManual_select').val();
                    d.listCheckValues = $("#valuesCheckElement").val();
                    if(d.vara == null && d.temperatura == null && d.temperaturaManual == null)
                        $("#filterHidden").val(false);
                    else
                        $("#filterHidden").val(true);
                    d.filter = $("#filterHidden").val();
                  }
            },
            "order": [[ 7, "desc" ]],
            "language": {
                "url": "//cdn.datatables.net/plug-ins/9dcbecd42ad/i18n/Portuguese-Brasil.json"
            },
            columnDefs: [
                { orderable: false, targets: [0,8,9] },
                {"className": "dt-center", "targets": "_all"}
            ],
            "pageLength":10,
            "serverSide": true,
            "responsive": true,
            "processing": true
        });
    }
}

function tableAnalista(create = true){
    if(create){
        var table = $('#dataTables3').DataTable({
            "ajax": {
                "url": $('#dataTables3').attr("ajaxpath"),
                "type": "POST"
            },
            "language": {
                "url": "//cdn.datatables.net/plug-ins/9dcbecd42ad/i18n/Portuguese-Brasil.json"
            },
            "order": [[ 4, "desc" ]],
            columnDefs: [
                { orderable: false, targets: [0] },
                {"className": "dt-center", "targets": "_all"}
            ],
            "pageLength":10,
            "serverSide": true,
            "responsive": true,
            "processing": true,
            "drawCallback": function(settings) {
                $( "button[input-send='modalProcesso']" ).click(function() {
                    let hash = $(this).attr("data-hash");
                    modaledicao(hash,"analista");
                });

            }, 
        })
    }else{
        $('#dataTables3').DataTable().ajax.reload();
    }
    cardsAnalise("analista");
}

function tableAdm(create = true){
    if(create){
        var table = $('#dataTables1').DataTable({
            "ajax": {
                "url": $('#dataTables1').attr("ajaxpath"),
                "type": "POST"
            },
            "language": {
                "url": "//cdn.datatables.net/plug-ins/9dcbecd42ad/i18n/Portuguese-Brasil.json"
            },
            "order": [[ 5, "desc" ]],
            columnDefs: [
                { orderable: false, targets: [0] },
                {"className": "dt-center", "targets": "_all"}
            ],
            "pageLength":10,
            "serverSide": true,
            "responsive": true,
            "processing": true,
            "drawCallback": function(settings) {
                $( "button[input-send='modalProcesso']" ).click(function() {
                    let hash = $(this).attr("data-hash");
                    modaledicao(hash,"adm");
                });
             },
        });
    }else{
        $('#dataTables1').DataTable().ajax.reload();
    }
    cardsAnalise("adm");
}

function cardsAnalise(type){
    if(type == "analista"){
        $.ajax({
            method: "POST",
            url: location.origin+"/pages/processos.php?status=ajax&case=17",
            async: true,
            cache: false,
            success: function (msg) {
                msg = JSON.parse(msg);
                if(msg.cards){
                    $("#cardsProcessosAnalista")[0].style.display = '';
                    if(msg.data.status == 200){
                        $("#processWait").text(msg.data.waitAnalise + " processos aguardando análise");
                        $("#processAnaliseNow").text(msg.data.inAnalise + " processos sendo analisados");
                        $("#processAnaliseToday").text(msg.data.allFinishToday + " processos analisados hoje");

                    }
                }
            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data'
        });
    }else if(type == "adm"){
        $.ajax({
            method: "POST",
            url: location.origin+"/pages/processos.php?status=ajax&case=18",
            async: true,
            cache: false,
            success: function (msg) {
                msg = JSON.parse(msg);
                $("#cardsProcessosAnalista")[0].style.display = '';
                if(msg.data.status == 200){
                    $("#processWait").text(msg.data.waitAnalise + " Processos Aguardando Validação");
                    $("#processAnaliseToday").text(msg.data.allFinishToday + " Processos Validados Hoje");

                }
            },
            error: function (error) {
                console.log(error);
            },
            contentType: false,
            processData: false,
            enctype: 'multipart/form-data'
        });
    }
}

function formatarMoeda(idInput) {
    var elemento = document.getElementById(idInput);
        var valor = elemento.value;
        if(valor.length > 1){

            valor = valor + '';
            valor = parseInt(valor.replace(/[\D]+/g, ''));
            valor = valor + '';
            valor = valor.replace(/([0-9]{2})$/g, ",$1");

            if (valor.length > 6) {
                valor = valor.replace(/([0-9]{3}),([0-9]{2}$)/g, ".$1,$2");
            }

            elemento.value = valor;
            if(valor == 'NaN') 
                elemento.value = '';
        }
}

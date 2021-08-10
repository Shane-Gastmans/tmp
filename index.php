<!DOCTYPE html>
<html>
   <head>
     <title>P&D bouw vorderingstaat</title>
     <meta name="viewport" content="width=device-width, initial-scale=1.0">
     <link rel="icon" type="image/png" href="/images/logo-small.jpg"/>
     <style>
       .selectedRow {
         border-top: 5px solid red!important;
         border-bottom: 5px solid red!important;
       }
       #excel_area td:hover {background-color:green;color:white;}
       .selectable {background-color:red;color:white;}
       .selected {background-color:green;color:white;}
       #fields {
          display: none;
         }
         #load_excel_form {
           position: sticky;
           top:15px;
           background-color: white;
           padding: 7px;
         }
         .container {
           max-height: 100%;
         }
         
         header {
           background-color: #c1001f;
           width: 100%;
           margin: 0;
           padding: 0;
         }

         header img {
           height: 90px;
         }


     </style>
     <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" />
   </head>
   <body>
     <header>
         <img src="/images/logo-small.jpg" alt="">
     </header>
     <div class="container">
     <form method="post" id="load_excel_form" enctype="multipart/form-data">
            <table id="filetable" class="table">
              <tr>
                <td width="25%" align="right">Selecteer Excel bestand</td>
                <td width="25%" align="center"><input type="file" name="select_excel" /></td>
		        <td width="25%" align="left"><input type="text" name="sheet" value="Offerte klant" /></td>
              </tr>
            </table>
            <input style="display:block;margin:auto;" type="submit" name="load" value="Upload excel bestand" width="100%" id="submitExcel" class="btn btn-primary" disabled />
              <div id="fields">
                  <label for="artnummer">Artnummer:</label>
                  <input type="text" id="artnummer" name="artnummer" class="selectable selected" value="" readonly>
                  <label for="eenheidsprijs">Beschrijving:</label>
                  <input type="text" id="beschrijving" name="beschrijving" class="selectable" value="" readonly>
                  <label for="eenheid">Eenheid:</label>
                  <input type="text" id="eenheid" name="eenheid" class="selectable" value="" readonly>
                  <!-- <<br> -->
                  <label for="eenheidsprijs">Eenheidsprijs:</label>
                  <input type="text" id="eenheidsprijs" name="eenheidsprijs" class="selectable" value="" readonly>
                  <label for="hoeveelheid">Hoeveelheid:</label>
                  <input type="text" id="hoeveelheid" name="hoeveelheid" class="selectable" value="" readonly>
                  <label for="totaal">Totaal:</label>
                  <input type="text" id="totaal" name="totaal" class="selectable" value="" readonly>
                  <input type="submit">
              </div>
	   </form>

      <br />
      <br />
      <div class="table-responsive">
       <span id="message"></span>
          
       <br />
          
        <div id="excel_area"></div>
      
      </div>
     </div>
    </div>
     <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js"></script>
  </body>
</html>
<script>
$(document).ready(function(){
  $('#load_excel_form').on('submit', function(event){
        event.preventDefault();
        $.ajax({
            url:"upload.php",
            method:"POST",
            data:new FormData(this),
            contentType:false,
            cache:false,
            processData:false,
            success:function(data)
            {
                if ($("#hoeveelheid").val() == "") {
                    $("#filetable").hide();
                    $("#submitExcel").hide();
                    $('#excel_area').html(data);
                    $('table').css('width','100%');
                    $("#excel_area td").on("click", function(){
                        $(".selected").val($(this).attr("class").split(" ")[0].match(/\d+/g) + "-" + $(this).closest("tr").attr("class").match(/\d+/g));
                        let next = $(".selected").next().next();
                        $(".selected").removeClass("selected");
                        next.addClass("selected");
                    });
                    $("#excel_area tr").on("click", function(){
                        console.log("clicked table row")
                        $(this).children().addClass("selectedRow");
                    });
                    $("#fields").show();
                  
                } else {
                    $("#message").html(data);
                }
            }
        })
    });

  $(".selectable").on("click", function(){
    $(".selected").removeClass("selected");
    $(this).addClass("selected");
  });
  
  $("#load_excel_form").on("change", function() {
      $("#submitExcel").attr("disabled", false)
  });

});

document.getElementById("load_excel_form").reset()

</script>

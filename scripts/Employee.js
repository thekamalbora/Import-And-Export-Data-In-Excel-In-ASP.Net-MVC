var GlobalURL = "http://localhost:5647/";


$(document).ready(function () {
    GetEmployeeRecord();
});


function EmployeeInsertUpdateData() {
    $.ajax({
        url: GlobalURL + "Home/EmployeeInsertUpdateData",
        type: "POST",
        data: {
            EmployeeID: $("#hfEmployeeID").val(),
            EmployeeName: $("#txtEmployeeName").val(),
            EmployeeAge: $("#txtEmployeeAge").val(),
            EmployeeProfile: $("#txtEmployeeProfile").val(),
            EmployeeSalary: $("#txtEmployeeSalary").val(),
            
        },
        success: function (data) {
            data = JSON.parse(data.d);
            if (data.Status == "1" || data.Status == "2") {
                alert(data.Result);
                GetEmployeeRecord();
                Reset();
            } else {
                alert(data.Result);
            }
            if (data.Focus != "") {
                $("#" + data.Focus).focus();
            }
        },
        error: function () {
            alert("Record not saved,Somethong wrong");
        }
    });
}


function GetEmployeeRecord() {
    $.ajax({
        url: GlobalURL + "Home/GetEmployeeRecord",
        type: "POST",
        data: {},
        success: function (data) {
            data = JSON.parse(data);
            $("#tbl").find("tr:gt(0)").remove();
            for (var i = 0; i < data.length; i++) {
                $("#tbl").append('<tr>    <td>' + data[i].EmployeeName + '</td>    <td>' + data[i].EmployeeAge + '</td>  <td>' + data[i].EmployeeProfile + '</td> <td>' + data[i].EmployeeSalary + '</td> </tr>');

            }
        },
        error: function () {
            alert("Record Not Load");
        }
    })
}

function Reset()
{
    $("#hfEmployeeID").val(0);
    $("#txtEmployeeName").val("");
    $("#txtEmployeeAge").val("");
    $("#txtEmployeeProfile").val("");
    $("#txtEmployeeSalary").val("");
}
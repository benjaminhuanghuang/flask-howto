Then router
```
from flask send_file, current_app
import xlsxwriter

def router_function():
  data = ""
  file_name = ""
  file_path = create_excel_file(file_name, data)
  return send_file(file_path)

```

```
def create_excel_file(file_name, data):
    temp_folder = os.path.join(current_app.static_folder, "temp/")
    file_path = os.path.join(temp_folder, file_name)

    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})

    # Add the worksheet data that the charts will refer to.
    rowIndex = 1
    for row in data:
        worksheet.write(rowIndex, 0, row["XX"])
   
    workbook.close()
    return file_path
```


Client

```
$scope.exportUserProgressData = function () {
            $scope.showWait();
            var params = $scope.getQueryParameter();
            $http({
                url: 'mathjoy/api/v1.0/exportUserProgressData',
                method: 'POST',
                responseType: 'arraybuffer',
                data: params,
                headers: {
                    'Content-type': 'application/json',
                    'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                }
            }).success(function (response) {
                $scope.hideWait();
                var blob = new Blob([response], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
                $scope.saveAs(blob, 'user_progress' + '.xlsx');
            }).error(function () {
                $scope.hideWait();
                //Some error log
            });
        };


 $scope.saveAs = function (blob, fileName) {
            if (window.navigator.msSaveOrOpenBlob) {
                navigator.msSaveBlob(blob, fileName);
            } else {
                var link = document.createElement('a');
                link.href = window.URL.createObjectURL(blob);
                link.download = fileName;
                link.click();
                window.URL.revokeObjectURL(link.href);
            }
        }
```
Add style
```
format = workbook.add_format({'bold': True, 'font_color': 'red'})
worksheet.write(rowIndex, 0, row["user_name"], format)

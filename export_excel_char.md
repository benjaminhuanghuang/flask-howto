```
from flask send_file, current_app
import xlsxwriter

def router_function():
    data = ""
    file_name = ""
    file_path = create_excel_file(file_name, data)
    return send_file(file_path)


def create_excel_file(file_name, data):
    date_grade_score_dict = data["dateGradeScore"]
    grade_list = data["gradeList"]
    temp_folder = os.path.join(current_app.static_folder, "temp/")
    file_path = os.path.join(temp_folder, file_name)

    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})

    # Add the worksheet data that the charts will refer to.
    worksheet.write(0, 0, 'Date')
    for index, grade in enumerate(grade_list):
        if int(grade) == 9:
            grade_label = "Algebra I"
        elif int(grade) == 10:
            grade_label = "Geometry"
        elif int(grade) == 11:
            grade_label = "Algebra II"
        elif int(grade) == 12:
            grade_label = "Pre-Calculus"
        else:
            grade_label = "Grade " + grade
        worksheet.write(0, index + 1, grade_label)

    dates = date_grade_score_dict.keys()
    dates.sort()

    current_row = 1
    for key in dates:
        worksheet.write(current_row, 0, key)

        grade_score = date_grade_score_dict[key]
        for index, grade in enumerate(grade_list):
            if grade in grade_score and grade_score[grade]["score"]:  # grade and score
                score = float(grade_score[grade]["score"])
            else:
                score = None

            worksheet.write(current_row, index + 1, score)
        current_row += 1

    # Create a new chart object. In this case an embedded chart.
    chart1 = workbook.add_chart({'type': 'line'})
    line_colors = [
        "#024ADD",
        "#00BEEB",
        "#549D28",
        "#F2CF03",
        "#F30100",
        "#9706FF",
        "#FF711B",
        "#8C0A91",
        "#3B5998",
        "#8C0C2F"
    ]
    for index, grade in enumerate(grade_list):
        line_color = line_colors[index % 10]
        chart1.add_series({
            'name': ['Sheet1', 0, index + 1],
            'categories': ['Sheet1', 1, 0, current_row, 0],
            'values': ['Sheet1', 1, index + 1, current_row, index + 1],
            'marker': {'type': 'diamond', 'size': 6, 'fill': {'color': line_color}, 'border': {'color': line_color}},
            'line': {'width': 3, 'color': line_color}
        })

    # 'gap'   # Blank data is shown as a gap. The default.
    # 'zero'  # Blank data is displayed as zero.
    # 'span'  # Blank data is connected with a line.
    chart1.show_blanks_as('span')

    # Add a chart title and some axis labels.
    chart1.set_size({'width': 800, 'height': 650})
    chart_name = "Progress Chart of " + data["userName"]
    chart1.set_title({'name': chart_name})
    chart1.set_x_axis({'name': 'Date'})
    chart1.set_y_axis({'name': 'Score', 'min': 0, 'max': 120})

    # Set an Excel chart style. Colors with white outline and shadow.
    chart1.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('H2', chart1, {'x_offset': 25, 'y_offset': 10})

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

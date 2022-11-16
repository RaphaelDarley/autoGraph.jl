import XLSX
import DataFrames
using Plots
using Printf
using Dates


function contains_string(vec, str)
    for item in vec
        if item == str
            return true
        end
    end
    return false
end

function run_graph_gen()

end


println("Welcome to autoGraph!")
println("This program is designed to automatically create graphs from the BP stats review")
print("Please enter path to XLSX file.\n>>>")
file_path = readline()
# println(file_path)
xf = XLSX.readxlsx(file_path)
worksheet_names = XLSX.sheetnames(xf)
worksheet_num = length(worksheet_names)
print("You have selected a file with $worksheet_num sheets, would you like them to be enumerated here?\n(y/N) ")
resp = readline()
if lowercase(resp) == "y"
    for name in worksheet_names
        println(name)
    end
end

sheet_name = ""
print("Please enter the name of the sheet to graph exactly\n>>>")
worksheet_name = readline()
while !contains_string(worksheet_names, worksheet_name)
    print("Worksheet not found in file, try again, or quit and select new file.\n>>>")
    global sheet_name = readline()
end

op_worksheet = xf[worksheet_name]

print("Is type(eg. country) in column A?\n(Y/n)")
type_col = "A"
resp = readline()
if lowercase(resp) == "n"
    print("Please enter type column.\n>>>")
    type_col = readline()
end

print("Please enter range of rows for types, currently only ranges are supported, please enter the starting row number followed by a colon and the ending number, the second number should be larger than the starting number eg. 5:30.\n>>>")
resp = readline()
type_range = split(resp, ":")
type_start = parse(Int, type_range[1])
type_end = parse(Int, type_range[2])

# println("$type_col$type_start:$type_col$type_end")

print("Is the x-axis(eg. year) in row 3?\n(Y/n)")
x_row = "3"
resp = readline()
if lowercase(resp) == "n"
    print("Please enter x-axis row.\n>>>")
    x_row = readline()
end

print("Please enter x-axis range in a similar manner to the type range eg. B:AC.\n>>>")
resp = readline()
x_range = split(resp, ":")
x_start = x_range[1]
x_end = x_range[2]

print("Please enter a path to save the files.\n>>>")
out_path = readline()
out_path = Base.Filesystem.mkpath(out_path)
# println("path created")

# println("start: $type_start")
# println("end: $type_end")

x_series = op_worksheet["$x_start$x_row:$x_end$x_row"][:]

for i in type_start:type_end
    type_name = op_worksheet["$type_col$i"]
    if (typeof(type_name) == Missing)
        # println("skipped")
        continue
    end
    type_name = strip(type_name)

    plot_name = "$worksheet_name-$type_name"

    series = op_worksheet["$x_start$i:$x_end$i"]

    plot_draft = plot(x_series, series[:], legend=false)
    title!(plot_draft, plot_name)

    png(plot_draft, "$out_path/$plot_name")
end
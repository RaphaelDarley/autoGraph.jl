import XLSX
using Plots
# using Printf
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

print("Please enter the name of the sheet to graph exactly\n>>>")
worksheet_name1 = readline()
while !contains_string(worksheet_names, worksheet_name1)
    print("Worksheet not found in file, try again, or quit and select new file.\n>>>")
    global worksheet_name1 = readline()
end

op_worksheet1 = xf[worksheet_name1]

type_col = "A"
x_row = "3"


print("Please enter range of rows for types, currently only ranges are supported, please enter the starting row number followed by a colon and the ending number, the second number should be larger than the starting number eg. 5:30.\n>>>")
resp = readline()
type_range = split(resp, ":")
type_start1 = parse(Int, type_range[1])
type_end1 = parse(Int, type_range[2])

print("Please enter x-axis range in a similar manner to the type range eg. B:AC.\n>>>")
resp = readline()
x_range = split(resp, ":")
x_start1 = x_range[1]
x_end1 = x_range[2]


# second sheet


print("Please enter the name of the second sheet to graph exactly\n>>>")
worksheet_name2 = readline()
while !contains_string(worksheet_names, worksheet_name2)
    print("Worksheet not found in file, try again, or quit and select new file.\n>>>")
    global worksheet_name2 = readline()
end

op_worksheet2 = xf[worksheet_name2]

print("Please enter range of rows for types, currently only ranges are supported, please enter the starting row number followed by a colon and the ending number, the second number should be larger than the starting number eg. 5:30.\n>>>")
resp = readline()
type_range = split(resp, ":")
type_start2 = parse(Int, type_range[1])
type_end2 = parse(Int, type_range[2])

print("Please enter x-axis range in a similar manner to the type range eg. B:AC.\n>>>")
resp = readline()
x_range = split(resp, ":")
x_start2 = x_range[1]
x_end2 = x_range[2]

print("Please enter a path to save the files.\n>>>")
out_path = readline()
out_path = Base.Filesystem.mkpath(out_path)
# println("path created")

# println("start: $type_start")
# println("end: $type_end")

x_series1 = op_worksheet1["$x_start1$x_row:$x_end1$x_row"][:]
x_series2 = op_worksheet2["$x_start2$x_row:$x_end2$x_row"][:]

# println("series 1: $x_series1")
# println("series 2: $x_series2")


if x_series1 != x_series2
    error("X series do not match")
else
    x_series = x_series1
end

data_series1 = Dict()

for i in type_start1:type_end1
    type_name = op_worksheet1["$type_col$i"]
    if (typeof(type_name) == Missing)
        continue
    end
    type_name = strip(type_name)

    series = op_worksheet1["$x_start1$i:$x_end1$i"]

    data_series1[type_name] = series
end


data_series2 = Dict()

for i in type_start2:type_end2
    type_name = op_worksheet2["$type_col$i"]
    if (typeof(type_name) == Missing)
        continue
    end
    type_name = strip(type_name)

    series = op_worksheet2["$x_start1$i:$x_end1$i"]

    data_series2[type_name] = series
end

# println(data_series1)
# println(data_series2)

for (type_name, series1) in data_series1
    if !haskey(data_series2, type_name)
        continue
    end

    plot_name = "$type_name-$worksheet_name1-$worksheet_name2"

    series1 = series1[:]
    series2 = data_series2[type_name][:]

    diff_series = [d1 - d2 for (d1, d2) in zip(series1, series2)]


    plot_draft = plot(x_series, series1, label=worksheet_name1, legend=:topleft)
    plot!(plot_draft, x_series, series2, label=worksheet_name2)
    plot!(plot_draft, x_series, diff_series, label="($worksheet_name1) - ($worksheet_name2)")

    if minimum(diff_series) < 0
        plot!(plot_draft, (x) -> 0, minimum(x_series), maximum(x_series), label="", line=(:black))
    end

    title!(plot_draft, plot_name)

    png(plot_draft, "$out_path/$plot_name")


end




# for i in type_start:type_end
#     type_name = op_worksheet["$type_col$i"]
#     if (typeof(type_name) == Missing)
#         # println("skipped")
#         continue
#     end
#     type_name = strip(type_name)

#     plot_name = "$worksheet_name-$type_name"

#     series = op_worksheet["$x_start$i:$x_end$i"]

#     plot_draft = plot(x_series, series[:], legend=false)
#     title!(plot_draft, plot_name)

#     png(plot_draft, "$out_path/$plot_name")
# end
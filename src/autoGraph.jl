import XLSX
using Plots
# using Printf
# using Dates
using TOML


function contains_string(vec, str)
    for item in vec
        if item == str
            return true
        end
    end
    return false
end



function run_graph_gen()

    worksheet_name = ""
    print("Please enter the name of the sheet to graph exactly\n>>>")
    worksheet_name = readline()
    while !contains_string(worksheet_names, worksheet_name)
        print("Worksheet not found in file, try again, or quit and select new file.\n>>>")
        global worksheet_name = readline()
    end

    op_worksheet = xf[worksheet_name]

    # print("Is type(eg. country) in column A?\n(Y/n)")
    type_col = "A"
    # resp = readline()
    # if lowercase(resp) == "n"
    #     print("Please enter type column.\n>>>")
    #     type_col = readline()
    # end

    print("Please enter range of rows for types, currently only ranges are supported, please enter the starting row number followed by a colon and the ending number, the second number should be larger than the starting number eg. 5:30.\n>>>")
    resp = readline()
    type_range = split(resp, ":")
    type_start = parse(Int, type_range[1])
    type_end = parse(Int, type_range[2])

    # println("$type_col$type_start:$type_col$type_end")

    # print("Is the x-axis(eg. year) in row 3?\n(Y/n)")
    x_row = "3"
    # resp = readline()
    # if lowercase(resp) == "n"
    #     print("Please enter x-axis row.\n>>>")
    #     x_row = readline()
    # end

    print("Please enter x-axis range in a similar manner to the type range eg. B:AC.\n>>>")
    resp = readline()
    x_range = split(resp, ":")
    x_start = x_range[1]
    x_end = x_range[2]

    println("1: Standard chart")
    println("2: Hubbert linearisation")
    println("3: Annual against cumulative")
    println("4: Smart HL")
    print("please enter enter number(s) of desired charts in comma seperated list eg. \"1,2\" default is standard\n>>>")
    type_resp = readline()
    if type_resp == ""
        chart_types = ["1"]
    else
        chart_types = [strip(t) for t in split(type_resp, ",")]
    end

    # println("chart types: $chart_types")

    println("0: none")
    println("1: Kb/d")

    print("Select source unit setting, leave blank for other.\n>>>")
    unit_setting = readline()
    if unit_setting == ""
        unit_setting = "0"
    end

    unit_setting = parse(Int, unit_setting)




    if "1" in chart_types # STANDARD GRAPH

        print("Please enter a path to save the standard files.\n>>>")
        out_path = readline()
        out_path = Base.Filesystem.mkpath(out_path)

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

            if unit_setting == 1
                series = 365 / 1000 .* series
            end


            plot_draft = plot(x_series, series[:], legend=false)
            title!(plot_draft, plot_name)

            if unit_setting == 1
                plot!(plot_draft, ylabel="annual Mb/a")
            end

            plot_name = join(split(plot_name, ":"))

            png(plot_draft, "$out_path/$plot_name")
        end
    end

    if "2" in chart_types # Hubbert linearisation

        print("Please enter a path to save the HL files.\n>>>")
        out_path = readline()
        out_path = Base.Filesystem.mkpath(out_path)

        for i in type_start:type_end
            type_name = op_worksheet["$type_col$i"]
            if (typeof(type_name) == Missing)
                # println("skipped")
                continue
            end
            type_name = strip(type_name)

            plot_name = "HL $worksheet_name-$type_name"

            series = op_worksheet["$x_start$i:$x_end$i"][:]

            if unit_setting == 1
                series = 365 / 1_000_000 .* series
            end


            cumulative::Vector{Float64} = [series[1]]
            apcp::Vector{Float64} = [1.0]

            for d in series
                push!(cumulative, cumulative[end] + d)
                push!(apcp, d / cumulative[end])
            end


            plot_draft = scatter(cumulative[2:end], apcp[2:end], legend=false)
            title!(plot_draft, plot_name)

            if unit_setting == 0
                plot!(plot_draft, xlabel="cumulative", ylabel="annual/cumulative")

            elseif unit_setting == 1
                plot!(plot_draft, xlabel="cumulative, Gb", ylabel="annual/cumulative")
            end

            plot_name = join(split(plot_name, ":"))
            png(plot_draft, "$out_path/$plot_name")
        end
    end

    if "3" in chart_types # Annual against cumulative

        print("Please enter a path to save the annaul vs cumulative files.\n>>>")
        out_path = readline()
        out_path = Base.Filesystem.mkpath(out_path)

        for i in type_start:type_end
            type_name = op_worksheet["$type_col$i"]
            if (typeof(type_name) == Missing)
                # println("skipped")
                continue
            end
            type_name = strip(type_name)

            plot_name = "$worksheet_name-$type_name"

            series = op_worksheet["$x_start$i:$x_end$i"][:]

            if unit_setting == 1
                series = 365 / 1000 .* series
            end

            cumulative::Vector{Float64} = [series[1]]

            for d in series
                push!(cumulative, cumulative[end] + d)
            end

            if unit_setting == 1
                cumulative = 1 / 1000 .* cumulative
            end

            plot_draft = scatter(cumulative[2:end], series[2:end], legend=false)
            title!(plot_draft, plot_name)

            if unit_setting == 0
                plot!(plot_draft, xlabel="cumulative", ylabel="annual")

            elseif unit_setting == 1
                plot!(plot_draft, xlabel="cumulative Gb", ylabel="annual Mb")
            end

            plot_name = join(split(plot_name, ":"))
            png(plot_draft, "$out_path/$plot_name")
        end
    end

    if "4" in chart_types

        print("Please enter a path to save the smart HL.\n>>>")
        out_path = readline()
        out_path = Base.Filesystem.mkpath(out_path)

        for i in type_start:type_end
            type_name = op_worksheet["$type_col$i"]
            if (typeof(type_name) == Missing)
                # println("skipped")
                continue
            end
            type_name = strip(type_name)

            plot_name = "HL $worksheet_name-$type_name"

            series = op_worksheet["$x_start$i:$x_end$i"][:]

            if unit_setting == 1
                series = 365 / 1_000_000 .* series
            end


            cumulative::Vector{Float64} = [series[1]]
            apcp::Vector{Float64} = [1.0]

            for d in series
                push!(cumulative, cumulative[end] + d)
                push!(apcp, d / cumulative[end])
            end

            series_start = 1

            grad_acc = []
            for i in eachindex(cumulative[1:end-1])
                Δy = apcp[i+1] - apcp[i]
                Δx = cumulative[i+1] - cumulative[i]

                grad = Δy / Δx
                push!(grad_acc, grad)
            end

            map!((m) -> m > -0.0025, grad_acc, grad_acc)

            for i in eachindex(grad_acc[1:end-1])
                if grad_acc[i] && grad_acc[i+1]
                    series_start = i
                    break
                end
            end


            plot_draft = scatter(cumulative[series_start:end], apcp[series_start:end], legend=false)
            title!(plot_draft, plot_name)

            if unit_setting == 0
                plot!(plot_draft, xlabel="cumulative", ylabel="annual/cumulative")

            elseif unit_setting == 1
                plot!(plot_draft, xlabel="cumulative, Gb", ylabel="annual/cumulative")
            end

            best_fit = lm(@formula(Y ~ X), DataFrame(X=cumulative, Y=apcp))
            fit_start = 1

            for i in eachindex(grad_acc[1:end-12])
                data = DataFrame(X=cumulative[i:end], Y=apcp[i:end])
                ols = lm(@formula(Y ~ X), data)

                if r2(ols) > r2(best_fit)
                    best_fit = ols
                    fit_start = i
                end
            end


            Qmax = -coef(best_fit)[1] / coef(best_fit)[2]


            if cumulative[fit_start] < Qmax
                plot!(plot_draft, (x) -> coef(best_fit)[1] + coef(best_fit)[2] * x, cumulative[fit_start], Qmax)
            else
                plot!(plot_draft, (x) -> coef(best_fit)[1] + coef(best_fit)[2] * x)
            end

            plot_name = join(split(plot_name, ":"))

            eurr_text = @sprintf "EURR: %.2f" Qmax

            if unit_setting == 1
                eurr_text *= "Gb"
            end

            if Qmax !== NaN
                annotate!(plot_draft, Qmax * 0.95, maximum(apcp[series_start:end]) * 0.9, text(eurr_text, 10, :black, :right, :bottom))
            end
            png(plot_draft, "$out_path/$plot_name")
        end
    end

end





### MAIN START

println("Welcome to autoGraph!")
println("This program is designed to automatically create graphs from the BP stats review")

config = Dict()
try
    global config = TOML.tryparsefile("config.toml")
catch
end

if isa(config, TOML.ParserError)
    config = config.table
end

if "XLSX-file" in keys(config)
    file_path = config["XLSX-file"]
    println("Using file from config.toml")
else
    print("Please enter path to XLSX file.\n>>>")
    file_path = readline()
end

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

global continue_prog = true

while continue_prog
    run_graph_gen()
    print("Would you like to graph another sheet from the same file?\n(N/y)")
    continue_resp = readline()
    if lowercase(continue_resp) == "y"
        global continue_prog = true
    else
        global continue_prog = false
    end
end


import XLSX
using Plots


xf = XLSX.readxlsx("bp-stats-review-2022-all-data.xlsx")
worksheet_name = "Oil Production - Barrels"
op_worksheet = xf[worksheet_name]

type_col = "A"
type_start = 5
type_end = 80
x_row = "3"
x_start = "B"
x_end = "BF"

unit_setting = 1

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
    for (i, p) in enumerate(apcp)
        if p < 0.06
            series_start = i
            break
        end
    end




    global plot_draft = scatter(cumulative[series_start:end], apcp[series_start:end], legend=false)
    title!(plot_draft, plot_name)

    if unit_setting == 0
        plot!(plot_draft, xlabel="cumulative", ylabel="annual/cumulative")

    elseif unit_setting == 1
        plot!(plot_draft, xlabel="cumulative, Gb", ylabel="annual/cumulative")
    end

    plot_name = join(split(plot_name, ":"))
    display(plot_draft)
end

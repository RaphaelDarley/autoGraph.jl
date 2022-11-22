import XLSX
using Plots
using DataFrames, GLM
using Printf
using Dates


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

# for i in type_start:type_end
# println(i)
i = 49 # Algeria
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
# for (i, p) in enumerate(apcp)
#     if p < 0.06
#         series_start = i
#         break
#     end
# end

grad_acc = []
for i in eachindex(cumulative[1:end-1])
    # println(i)
    Δy = apcp[i+1] - apcp[i]
    Δx = cumulative[i+1] - cumulative[i]

    grad = Δy / Δx
    # println(Δy / Δx) 
    push!(grad_acc, grad)
end

map!((m) -> m > -0.0025, grad_acc, grad_acc)

for i in eachindex(grad_acc[1:end-1])
    if grad_acc[i] && grad_acc[i+1]
        series_start = i
        break
    end
end


# for i in reverse(eachindex(grad_acc[1: end-12]))
#     # println(i)
#     grad_diff = (grad_acc[i+1] - grad_acc[i]) / grad_acc[i + 1]

#     println(grad_diff)
# end


global plot_draft = scatter(cumulative[series_start:end], apcp[series_start:end], legend=false)
title!(plot_draft, plot_name)

if unit_setting == 0
    plot!(plot_draft, xlabel="cumulative", ylabel="annual/cumulative")

elseif unit_setting == 1
    plot!(plot_draft, xlabel="cumulative, Gb", ylabel="annual/cumulative")
end

global best_fit = lm(@formula(Y ~ X), DataFrame(X=cumulative, Y=apcp))
global fit_start = 1

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

# if Qmax !== NaN && cumulative[end] < Qmax
#     annotate!(plot_draft, Qmax, apcp[series_start], text(eurr_text, 10, :black, :right, :top))
# end

# watermark = @sprintf "generated on: %s, by autoGraph (Raphael Darley)\n from: %s" Date(now()) basename(file_path)

# annotate!(plot_draft, cumulative[end] * 0.35, apcp[series_start], text(watermark, 6, :black, :left, :bottom))


watermark = @sprintf "generated on: %s, by autoGraph (Raphael Darley)\n from: %s" Date(now()) basename(file_path)

if Qmax !== NaN && cumulative[end] < Qmax
    annotate!(plot_draft, Qmax, maximum(apcp[series_start:end]), text(eurr_text, 10, :black, :right, :top))
    annotate!(plot_draft, Qmax * 0.6, maximum(apcp[series_start:end]), text(watermark, 6, :black, :centre, :bottom))
else
    annotate!(plot_draft, cumulative[end] * 0.6, maximum(apcp[series_start:end]), text(watermark, 6, :black, :centre, :bottom))
end





display(plot_draft)
# end

display(plot_draft)
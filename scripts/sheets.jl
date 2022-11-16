import XLSX
import DataFrames
using Plots
using Printf
using Dates

xf = XLSX.readxlsx("A:/Raphael/projects/Graphs/bp-stats-review-2022-all-data.xlsx")

oilProd_barrels = xf["Oil Production - Barrels"]

for i in 5:80
    country_name = oilProd_barrels["A$i"]
    if (typeof(country_name) == Missing)
        println("skipped")
        continue
    end

    country_name = strip(country_name)

    prod_data = oilProd_barrels["B$i:BF$i"]

    country_plot = plot(oilProd_barrels["B3:BF3"][:], prod_data[:], title=country_name, legend=false)

    annotation = @sprintf "generated on : %s, by autoGraph from bp review 2022 (Raphael Darley)" Date(now())
    annotate!(country_plot, 0, 1, text(annotation, :left, 10))

    png(country_plot, "./output/oil-barrels/$country_name-oil-barrels")

end

gas_Bcf = xf["Gas Production - Bcf"]

for i in 5:78
    country_name = gas_Bcf["A$i"]
    if (typeof(country_name) == Missing)
        println("skipped")
        continue
    end
    country_name = strip(country_name)


    prod_data = gas_Bcf["B$i:BA$i"]

    country_plot = plot(gas_Bcf["B3:BA3"][:], prod_data[:], title=country_name, legend=false)
    title!(country_plot, "$country_name-gas_bcf")

    # annotation = @sprintf "generated on : %s, by autoGraph from bp review 2022 (Raphael Darley)" Date(now())
    # annotate!(country_plot, 0, 1, text(annotation, :left, 10))

    png(country_plot, "./output/gas-bcf/$country_name-gas_bcf")

end
using XLSX
using TOML

config = TOML.parsefile("config.toml")

if !("op_dir" in keys(config))
    error("No operational directory found in config.toml")
else
    op_dir = config["op_dir"]
end

files = Base.Filesystem.readdir(op_dir; join=true)

filter!(endswith(".xlsx"), files)

XLSX.openxlsx(joinpath(op_dir, "merged.xlsx"), mode="w") do xf_acc
    source_book_acc = []
    for file_path in files
        println(file_path)
        xf = XLSX.readxlsx(file_path)
        if length(XLSX.sheetnames(xf)) > 1
            error("only one worksheet per file is currently supported")
        end

        sheet_name = basename(file_path)[1:end-5]

        XLSX.addsheet!(xf_acc, sheet_name)
        xf_acc[sheet_name]["A1"] = xf[1][:]
    end
end
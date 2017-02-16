# Supplier monitoring (test on Qunsheng)

# filter IQC df for Qunsheng 

sup <- "S801805"

IQC_sup <- IQC_df %>%
        filter(CardCode == sup)

cbPalette <- c("#999999", "#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00", "#CC79A7")

gradientcolour <- c("lightgreen", "grey", "yellow", "orange", "red")

ggplot(IQC_sup, aes(x = InspDate, y = Result, colour = Severity)) +
        geom_point(size = 7, alpha = 0.8) +
        scale_colour_gradientn(colours = gradientcolour)  +
        facet_wrap(~ItemCode) +
        ggtitle(paste(IQC_df$CardFName[IQC_df$CardCode == sup][1], "Incoming Quality Performance"))

ggplot(IQC_sup, aes(x = InspDate, y = Result, colour = Severity)) +
        geom_point(size = 7, alpha = 0.7) +
        scale_colour_gradient(colours = "red")  +
        facet_wrap(~ItemCode) +
        ggtitle(paste(IQC_df$CardFName[IQC_df$CardCode == sup][1], "Incoming Quality Performance"))

######
# export result and PPAP list in csv
mainDir <- "C:/Users/PB/SkyDrive/DG Evolt/QAS_Data"
subDir <- "Qunsheng_followup"

dir.create(file.path(mainDir, subDir), showWarnings = FALSE)

setwd(file.path(mainDir, subDir))
File_Name <- paste("IQC_", "ThisWeek", "Qunsheng_followup", sep = "")
write.csv(IQC_PPAP, File_Name)

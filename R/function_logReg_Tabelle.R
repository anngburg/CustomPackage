
function_logReg_Tabelle <- function(model, data, dv) {
    logReg_fit <- stats::anova(model, test = "Chisq")

    lr.chi <- model$null.deviance - model$deviance
    lr.df  <- model$df.null       - model$df.residual
    lr.p   <- stats::pchisq(lr.chi, lr.df, lower.tail = FALSE)
    overall_fit <- data.frame(Chi2 = lr.chi,
                              df   = lr.df,
                              p    = lr.p)

    logReg_model <- summary(model)
    logReg_ORs <- as.data.frame(exp(stats::coef(model))) #ORs
    logReg_Beta <- as.data.frame(stats::coef(model)) #Beta-Gewichte
    logReg_coefficients <- as.data.frame(logReg_model$coefficients)
    logReg_CI <- as.data.frame(stats::confint(model))
    log_Reg_Wald <- ((logReg_coefficients$Estimate)^2 / (logReg_coefficients$`Std. Error`))

    Tabelle_logReg <- cbind(logReg_Beta, logReg_ORs, logReg_coefficients$Estimate, logReg_coefficients$`Std. Error`, logReg_coefficients$`z value`,
                            logReg_coefficients$`Pr(>|z|)`, logReg_CI$`2.5 %`, logReg_CI$`97.5 %`, log_Reg_Wald)
    colnames(Tabelle_logReg) <- c("Beta-Gewichte", "Odds Ratios", "Estimate", "Standardfehler", "z-Statistik", "p-Wert", "2.5 %", "97.5 %", "Chi² (Wald)")
    Tabelle_logReg <- as.data.frame(Tabelle_logReg)

    Tabelle_logReg <- Tabelle_logReg %>% dplyr::mutate_if(is.numeric, round, digits = 3)
    Tabelle_logReg$`95%iges KI` <- paste("[", Tabelle_logReg$`2.5 %`, ";", Tabelle_logReg$`97.5 %`, "]")
    Tabelle_logReg <- dplyr::select(Tabelle_logReg, c(`Odds Ratios`, `Beta-Gewichte`, `95%iges KI`, Standardfehler, Estimate, `z-Statistik`, `Chi² (Wald)`, `p-Wert`,))

    prob_logReg <- stats::predict(model, type = "response")
    roc_curve <- pROC::roc(data[[dv]], prob_logReg) #ROC Kurve berechnen


    auc_value <- pROC::auc(roc_curve) #AUC Wert extrahieren
    auc_CI <- as.data.frame(pROC::ci.auc(roc_curve)) %>% dplyr::mutate_if(is.numeric, round, digits = 2)
    auc_CIlower <- auc_CI[1, ]
    auc_CIupper <- auc_CI[3, ]
    auc_CI <- paste("[",auc_CIlower,";",auc_CIupper,"]")
    print(auc_value)

    predictedKontakt <- ifelse(prob_logReg > 0.5, 1, 0)
    actualKontakt <- data[[dv]]

    conf_matrix <- caret::confusionMatrix(as.factor(predictedKontakt), as.factor(actualKontakt))
    Sensitivität <- conf_matrix$byClass["Sensitivity"]
    Spezifität <- conf_matrix$byClass["Specificity"]
    ProzentRichtige <- conf_matrix$overall["Accuracy"]

    NagelkerksR <- fmsb::NagelkerkeR2(model)
    AIC <- AIC(model)
    BIC <- BIC(model)

    ModellGüte <- cbind(auc_value, auc_CI, Sensitivität, Spezifität, ProzentRichtige, NagelkerksR$R2, AIC, BIC)
    rownames(ModellGüte) <- c("Werte")
    colnames(ModellGüte) <- c("AUC", "95% KI (AUC)", "Sensitivität", "Spezifität", "Prozentsatz der Richtigen", "Nagelkerkes R²", "AIC", "BIC")
    ModellGüte <- as.data.frame(ModellGüte)

    for(colname in colnames(ModellGüte)) {
        if (colname != "95% KI (AUC)") {
            ModellGüte[[colname]] <- as.numeric(ModellGüte[[colname]])
        }
    }

    ModellGüte <- ModellGüte %>% mutate_if(is.numeric, round, digits = 2)

    model_name <- deparse(substitute(model))

    if (!is.character(model_name)) {
        stop("The 'model_name' variable must be a character string.")
    }

    filename <- sprintf("%s_Tabelle.xlsx", model_name)

    xlsx::write.xlsx(Tabelle_logReg, file = filename, sheetName = "log Modell")
    xlsx::write.xlsx(ModellGüte, file = filename, sheetName = "Modell Güte", append = TRUE)
    xlsx::write.xlsx(overall_fit, file = filename, sheetName = "Modell Fit", append = TRUE)
}

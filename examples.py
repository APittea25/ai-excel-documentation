# # examples.py

# # Purpose section few-shot example
# purpose_example = (
#     "Example:\n"
#     "This model projects annuity cashflows under stochastic mortality assumptions "
#     "using the Lee-Carter framework. It incorporates longevity improvement trends "
#     "and produces annuity tables.\n\n"
# )

# # Input description example
# input_example = (
#     "Example:\n"
#     "- Historical mortality rates (log-transformed) by age and year.\n"
#     "- Initial estimates of the kappa_t time index for mortality trends.\n"
#     "- beta_x values representing the average age-specific log mortality.\n"
# )

# # Output description example
# output_example = (
#     "Example:\n"
#     "Present value of future annuity payments\n\n"
# )

# # Logic step example
# logic_example = (
#     "Example:\n"
#     "_c1_: Calculate log mortality rates using base qx inputs.\n"
#     "_c2_: Apply Lee-Carter formula: log(m_x,t) = α_x + β_x * κ_t.\n"
#     "_c3_: Exponentiate to recover projected qx values.\n"
#     "_c4_: Adjust mortality for cohort or shock effects, if applicable.\n"
# )

# # Assumption example
# assumption_example = (
#     "Example:\n"
#     "- Future mortality improvements are projected using a linear trend in the κ_t parameter.\n"
#     "- Age-specific parameters α_x and β_x are fixed over the projection horizon.\n"
# )

# # Check description example
# check_example = (
#     "Example:\n"
#     "Ensure present value is non-negative and increases with lower discount rate.\n"
# )

# examples.py

# Purpose section few-shot example
purpose_example = (
    "Example:\n"
    "This model projects annuity cashflows under stochastic mortality assumptions "
    "using the Lee-Carter framework. It incorporates longevity improvement trends "
    "and produces annuity tables.\n\n"
)

# Input description example
input_example = (
    "Example:\n"
    "- Discount rate for calculating present value of future benefits.\n"
    "- Policyholder age at entry used for eligibility rules.\n\n"
)

# Output description example
output_example = (
    "Example:\n"
    "Present value of future annuity payments, aggregated by cohort and discount scenario.\n\n"
)

# Logic step example
logic_example = (
    "Example:\n"
    "_c1_: Calculate base mortality using qx.\n"
    "_c2_: Apply improvement factors.\n"
)

# Assumption example
assumption_example = (
    "Example:\n"
    "- Mortality improvements are assumed to follow a constant 1.5% trend.\n"
    "- Discount rate fixed at 3% annually.\n"
)

# Check description example
check_example = (
    "Example:\n"
    "Ensure present value is non-negative and increases with lower discount rate.\n"
)

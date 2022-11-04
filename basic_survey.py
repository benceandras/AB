import openpyxl as xl
import numpy as np
import time

survey_workbook = xl.load_workbook('D:\Google Drive\himoney\himoney_survey_logic.xlsx')

input_sheet = survey_workbook["Test"]

input_investment_amount = input_sheet['B5']
input_time_horizon = input_sheet['B6']
input_risk_aversion = input_sheet['B7']

# Order of weights is: investment amount, time horizon, risk aversion
evaluation_weights = np.array([0.3, 0.5, 0.2], float)

# Currently, there are 3 threshold steps
thresholds = 3
minimum_score = 0.2
maximum_score = 0.8
threshold_step_score = (maximum_score - minimum_score) / thresholds
low_score = minimum_score + threshold_step_score
high_score = low_score + threshold_step_score


investment_amount_threshold1 = 5000000
investment_amount_threshold2 = 10000000
investment_amount_threshold3 = 20000000
investment_amount_threshold4 = 40000000

time_horizon_threshold1 = 1
time_horizon_threshold2 = 3
time_horizon_threshold3 = 5
time_horizon_threshold4 = 10

risk_aversion_correction_factor = 0.8

if input_investment_amount.value in range(0, investment_amount_threshold1):
    evaluate_input_investment_amount = minimum_score
elif input_investment_amount.value in range(investment_amount_threshold1, investment_amount_threshold2):
    evaluate_input_investment_amount = \
        minimum_score + threshold_step_score * (input_investment_amount.value / investment_amount_threshold2)
elif input_investment_amount.value in range(investment_amount_threshold2, investment_amount_threshold3):
    evaluate_input_investment_amount = \
        low_score + threshold_step_score * (input_investment_amount.value / investment_amount_threshold3)
elif input_investment_amount.value in range(investment_amount_threshold3, investment_amount_threshold4):
    evaluate_input_investment_amount = \
        high_score + threshold_step_score * (input_investment_amount.value / investment_amount_threshold4)
else:
    evaluate_input_investment_amount = maximum_score

if input_time_horizon.value in range(0, time_horizon_threshold1):
    evaluate_input_time_horizon = minimum_score
elif input_time_horizon.value in range(time_horizon_threshold1, time_horizon_threshold2):
    evaluate_input_time_horizon = \
        minimum_score + threshold_step_score * (input_time_horizon.value / time_horizon_threshold2)
elif input_time_horizon.value in range(time_horizon_threshold2, time_horizon_threshold3):
    evaluate_input_time_horizon = \
        low_score + threshold_step_score * (input_time_horizon.value / time_horizon_threshold3)
elif input_time_horizon.value in range(time_horizon_threshold3, time_horizon_threshold4):
    evaluate_input_time_horizon = \
        high_score + threshold_step_score * (input_time_horizon.value / time_horizon_threshold4)
else:
    evaluate_input_time_horizon = maximum_score

evaluate_input_risk_aversion = (input_risk_aversion.value / 100) * risk_aversion_correction_factor

evaluation_results = np.asarray(
    [evaluate_input_investment_amount, evaluate_input_time_horizon, evaluate_input_risk_aversion])

print(evaluation_results)

himoney_profile_score = np.dot(evaluation_results, evaluation_weights)
time_flag = time.strftime('%a, %d %b %Y %H:%M:%S %Z(%z)')

print(evaluation_results[:])

input_sheet['A11'] = "Python script result as of " + time_flag
input_sheet['A12'] = "himoney profile score:"
input_sheet['B12'] = himoney_profile_score

input_sheet['A13'] = "Evaluation results were the following:"
input_sheet['B13'] = evaluation_results[0]
input_sheet['B14'] = evaluation_results[1]
input_sheet['B15'] = evaluation_results[2]

survey_workbook.save('D:\Google Drive\himoney\himoney_survey_logic.xlsx')

print(himoney_profile_score)
print(threshold_step_score)
print(evaluation_results[:])
# print(input_time_horizon.value, evaluate_input_time_horizon)

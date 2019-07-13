import pandas as pd
import numpy as np

# import data
ts = pd.read_excel('Female-Competition26-5-2019.xlsx', sheet_name = 'Transaction 26-5-2019')
trader_info = pd.read_excel('Female-Competition26-5-2019.xlsx', sheet_name = 'Quota 26-5-2019')

# activity
def activity(ts):
	activity = ts[ts.Type == 'female trade'].groupby('ID', sort = True)['Type'].agg({'activity':np.size})
	return activity

# activity used
def activity_used(ts):
	used = trader_info[['ID','Used']].set_index('ID')
	act = activity(ts)
	table_act = act.merge(used, how = 'left', right_index = True, left_index = True)
	table_act['activity_used'] = table_act.activity/table_act.Used
	table_act.drop(['activity','Used'],axis = 1, inplace = True)
	return table_act

# distance
def distance(ts):
	dist = ts['distance'] = abs(ts['Price Open']- ts['Price Close'])
	total_dist = ts.groupby('ID', sort = True)['distance'].agg({'distance':np.sum})
	return total_dist

# average distance
def avg_distance(ts):
	dist = ts['distance'] = abs(ts['Price Open']- ts['Price Close'])
	avg_dist = ts.groupby('ID', sort = True)['distance'].agg({'avg_distance':np.mean})
	return avg_dist

# profit
def profit(ts):
	profit = ts.groupby('ID', sort = True)['Balance'].agg({'profit':np.max})
	return profit

# profit used
def profit_used(ts):
	used = trader_info[['ID','Used']].set_index('ID')
	prof = profit(ts)
	table_prof = prof.merge(used, how = 'left', right_index = True, left_index = True)
	table_prof['profit_used'] = table_prof.profit/table_prof.Used
	table_prof.drop(['profit','Used'],axis = 1, inplace = True)
	return table_prof

# combine all results
def result():
	act = activity(ts)
	act_used = activity_used(ts)
	dist = distance(ts)
	avg_dist = avg_distance(ts)
	prof = profit(ts)
	prof_used = profit_used(ts)
	final = pd.concat([act,act_used,dist,avg_dist,prof,prof_used],axis = 1)
	return final

# convert to excel form
weekly_result = result()
trader_profile = trader_info[['Email', 'ID', 'Name', 'Lastname']].drop_duplicates().set_index('ID')
weekly_performance = trader_profile.merge(weekly_result,how = 'left', right_index = True, left_index = True)
weekly_performance = weekly_performance.fillna(0)
weekly_performance.to_excel('weekly_performance.xlsx')




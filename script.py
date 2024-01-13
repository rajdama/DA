import pandas as pd

def report(file):

    df = pd.read_excel(file) 
    chartReport = pd.DataFrame(columns=['Week','Incoming Chats', 'Unique Users','Closed By Bot','Bot Deflection %','Closed By Agents']).reset_index(drop=True)
    chartReport['Week'] = ["01 Feb - 07 Feb","08 Feb - 14 Feb","15 Feb - 21 Feb","22 Feb - 28 Feb"]

    df['AgentFirstResponseTime'] = df['AgentFirstResponseTime'].replace('-', pd.to_datetime(0, unit='s').time())

    ChatStartTime = pd.to_datetime(df['ChatStartTime'], errors='coerce')
    ChatEndTime = pd.to_datetime(df['ChatEndTime'], errors='coerce')

    def standardize_date(dt):
        if ((1<=dt.day) & (dt.day<=9)):
            month = dt.day
            dt = dt.replace(day=dt.month)
            dt = dt.replace(month=month)
        return dt

    df['ChatStartTime'] = ChatStartTime.apply(standardize_date)
    df['ChatEndTime'] = ChatEndTime.apply(standardize_date)
    days =  df['ChatStartTime'].dt.day

    weeks = [
        df[(1 <=days) &  (days <= 7)], df[(8 <=days) &  (days <= 14)], 
        df[(15 <=days) &  (days <= 21)], df[(22 <=days) &  (days <= 28)] 
    ]

    def applyChartFilter(week):
        incomingChats = len(week["RoomCode"])
        userid =  len(pd.unique(week["UserId"]))
        closedByBot = len(week[week['ClosedBy'].isin(["System"])])
        botDeflection = round(closedByBot/incomingChats,4)*100
        closedByAgents = len(week.loc[week['ClosedBy'] != 'System']) 
        return [incomingChats, userid, closedByAgents, botDeflection, closedByAgents]
    
    for i, week in enumerate(weeks, start=0):
        weekData = applyChartFilter(week)
        for j in range(1,6):
            chartReport.iloc[i,j] = weekData[j-1]

    agentReport = pd.DataFrame(columns=['Week','Agent Name','Chats Resolved', 'Avg Agent First Response Time (seconds)','Avg Agent Chat Resolution Time (seconds)'
                        ,'Average Agent CSAT Score','Business Hours CSAT','Outside Business Hours CSAT']).reset_index(drop=True)
    agentReport['Week'] = ["01 Feb - 07 Feb","08 Feb - 14 Feb","15 Feb - 21 Feb","22 Feb - 28 Feb"]

    for i, week in enumerate(weeks, start=0):
        AgentNames = pd.unique(week.loc[week['ClosedBy'] != 'System']['ClosedBy'])
        agentReport.iat[i,1] = AgentNames
    
    agentReport =  agentReport.explode('Agent Name')

    def calculate_csat_score(df):
        positive_csat_counts = len(df[df['CsatScore'].isin([4, 5])])
        negative_csat_counts = len(df[df['CsatScore'].isin([1,2,3])])
        if positive_csat_counts != 0:
            avg_csat = positive_csat_counts / (positive_csat_counts + negative_csat_counts)
        else:
            avg_csat = 0
        return round(avg_csat, 2)

    def agentReportFilter(row):
        data = weeks[0] if row.iloc[0] == "01 Feb - 07 Feb" else weeks[1] if row.iloc[0] == "08 Feb - 14 Feb" else weeks[2] if row.iloc[0] == "15 Feb - 21 Feb" else weeks[3]
        data = data.loc[data['ClosedBy'] == row.iloc[1]]
        row.iloc[2] = len(data)
        sec_avg = round(data['AgentFirstResponseTime'].apply(lambda t: t.hour * 3600 + t.minute * 60 + t.second).mean(), 2)
        row.iloc[3] = sec_avg
        row.iloc[4] = round(((data['ChatEndTime'] - data['ChatStartTime']).dt.total_seconds()).mean(), 2)
        row.iloc[5] = calculate_csat_score(data)
        start_time = pd.to_datetime('10:00:00').time()
        end_time = pd.to_datetime('17:00:00').time()
        office_hours = data[(data['ChatStartTime'].dt.time >= start_time) & (data['ChatEndTime'].dt.time <= end_time)]
        row.iloc[6] = calculate_csat_score(office_hours)
        non_office_hours = data[(data['ChatStartTime'].dt.time < start_time) | (data['ChatEndTime'].dt.time > end_time)]
        row.iloc[7] = calculate_csat_score(non_office_hours)
        return row

    agentReport = agentReport.apply(agentReportFilter, axis=1)

    with pd.ExcelWriter('output_file.xlsx') as writer:
        # Write each DataFrame to a different sheet
        chartReport.to_excel(writer, sheet_name='Chart Report', index=False)
        agentReport.to_excel(writer, sheet_name='Agent Report', index=False)

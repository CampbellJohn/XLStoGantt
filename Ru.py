import Deliv
import Dra
import PP

""" Scrape the Calendar """
cal = Deliv.CalendarPull("Ref\Calendar_Jan22.xlsx")

""" Sort Into Quarters """
quarters = Deliv.Quarters(cal)

""" Create our PM objects """
HPC = Deliv.ProductManager(quarters, "Jay", "HPCS")
Uri = Deliv.ProductManager(quarters, "Kevin", "Urika")
SDM = Deliv.ProductManager(quarters, "Jason", "SDM")
Clu = Deliv.ProductManager(quarters, "Amar", "Clusters")
IBM = Deliv.ProductManager(quarters, "IBM", "IBM")

""" Creating a list of lists. All gantt lists moved to PM_Q1_gantt """
PMList = [HPC, Uri, SDM, Clu, IBM]
PM_Q1_gantt = []

for x in PMList:
    for y in range(len(x.Q1_gantt)):
        PM_Q1_gantt.append(x.Q1_gantt[y])

""" Create Powerpoint Presentation """
ppt = PP.Powerpoint()
ppt.gantt_preso(PM_Q1_gantt)

""" Save out .xls files """
for x in PMList:
    Deliv.save(x.Q1_xls, x.filename)

Camptest = Deliv.CampaignRanges(quarters)

testlist = Camptest.Q1_clist

# print(Camptest.Q1_sorted[3])

# print(Camptest.Q1_minmax)

testimg = Dra.Gantt(Camptest.Q1_minmax)

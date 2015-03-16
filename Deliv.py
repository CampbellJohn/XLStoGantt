import xlrd
import xlwt
import time
import datetime

""" To Do: 
    Determine campaign ranges. Create a list of rows that a gantt chart can be
    drawn from.
    Prepare a list of deliverables for Marcomm.
"""

""" SplitDeliverables XLS Key """
# row[0] = Stage
# row[1] = PM
# row[2] = Campaign
# row[3] = Product
# row[4] = Name
# row[5] = Due Day Integer
# row[6] = Duration
# row[7] = Due Date
# row[8] = iscomplete, "Y" or "N"
# row[9] = Which campaign sheet, 1-3
# row[10] = Which Quarter
# row[11] = Type of Asset

""" SplitDeliverables Marcomm Package"""
# row[0] = Stage
# row[1] = PM
# row[2] = Product
# row[3] = Name
# row[4] = Request Date
# row[5] = Due Date
# row[6] = Iscomplete, "Y" or "N"

class CalendarPull:
    def __init__(self, filename):
        self.filename = filename
        self.data = self.router()

    def router(self):
        pull = self.loadxls()
        assets = self.assetfinder(pull)
        split = self.split(assets)
        final = self.iscomplete(split)
        return final

    def loadxls(self):
        file_location = self.filename
        workbook = xlrd.open_workbook(file_location) # Open Workbook
        sheet = workbook.sheet_by_index(0) # First Sheet
        data = [[sheet.cell_value(r,c) for c in range(sheet.ncols)]           \
               for r in range(sheet.nrows)]
        return data

    def assetfinder(self, data):
        assets = []

        # Go through each column and grab assets.
        for x in range(len(data)):
            if "January" in data[x][1]: # 1, 3, 19
                pass
            elif "ebinar" in data[x][1]:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM 
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Prod
                                 data[x][1], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                 data[x][3], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10] Quarter
                                   "webinar" # [11] Type
                                ])
            else:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Prod
                                 data[x][1], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                 data[x][3], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                     "email" # [11] Type
                                ])

        for x in range(len(data)):
            if "February" in data[x][4]: # 4, 6, 19
                pass
            elif "ebinar" in data[x][4]:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Prod
                                 data[x][4], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                 data[x][6], # [7] DueDate 
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10] Quarter
                                   "webinar" # [11] Type
                                ])
            else:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Prod
                                 data[x][4], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                 data[x][6], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                     "email" # [11] Type
                                ])

        for x in range(len(data)):
            if "March" in data[x][7]: # 7, 9, 19
                pass
            elif "ebinar" in data[x][7]:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Prod
                                 data[x][7], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                 data[x][9], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                   "webinar" # [11] Type
                                ])
            else:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Prod
                                 data[x][7], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                 data[x][9], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                     "email" # [11] Type
                                ])

        for x in range(len(data)):
            if "April" in data[x][10]: # 10, 12, 19
                pass
            elif "ebinar" in data[x][10]:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Pro
                                data[x][10], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                data[x][12], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                  "webinar"  # [11] Type
                                ])
            else:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Pro
                                data[x][10], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                data[x][12], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                     "email" # [11] Type
                                ])
        return assets

        for x in range(len(data)):
            if "May" in data[x][13]: # 13, 15, 19
                pass
            elif "ebinar" in data[x][13]:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Pro
                                data[x][13], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                data[x][15], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                  "webinar"  # [11] Type
                                ])
            else:
                assets.append([          "", # [0] Stage
                                data[x][19], # [1] PM
                                 data[x][0], # [2] Campaign
                                 data[x][0], # [3] Pro
                                data[x][13], # [4] Name
                                         "", # [5] Due Int
                                         "", # [6] Duration
                                data[x][15], # [7] DueDate
                                        "N", # [8] Iscomp
                                         "", # [9] Sheet
                                         "", # [10]  Quarter
                                    "email"  # [11] Type
                                ])
                
        for x in range(len(data)):
            if "June" in data[x][16]: # 16, 18, 19
                pass
            elif "ebinar" in data[x][16]:
                assets.append([            "", # [0] Stage
                                  data[x][19], # [1] PM
                                   data[x][0], # [2] Campaign
                                   data[x][0], # [3] Pro
                                  data[x][16], # [4] Name
                                           "", # [5] Due Int
                                           "", # [6] Duration
                                  data[x][18], # [7] DueDate
                                          "N", # [8] Iscomp
                                           "", # [9] Sheet
                                           "", # [10]  Quarter
                                    "webinar"  # [11] Type
                                ])
            else:
                assets.append([            "", # [0] Stage
                                  data[x][19], # [1] PM
                                   data[x][0], # [2] Campaign
                                   data[x][0], # [3] Pro
                                  data[x][16], # [4] Name
                                           "", # [5] Due Int
                                           "", # [6] Duration
                                  data[x][18], # [7] DueDate
                                          "N", # [8] Iscomp
                                           "", # [9] Sheet
                                           "", # [10]  Quarter
                                      "email"  # [11] Type
                                ])

    def split(self, data):
        split = []
        yearNum = 42003
        for x in range(len(data)):
            if data[x][7]:
                # If a date exists, append rows of deliverables.
                split.append([    "PM Writes", # [0] Stage
                                   data[x][1], # [1] PM
                                   data[x][2], # [2] Campaign
                                   data[x][3], # [3] Prod
                                   data[x][4], # [4] Name
                    ((data[x][7]-yearNum)-20), # [5] Due Int
                                            7, # [6] Duration
                              (data[x][7]-20), # [7] DueDate
                                   data[x][8], # [8] IsComp
                                   data[x][9], # [9] Sheet
                                  data[x][10], # [10] Quarter
                                  data[x][11], # [11] Type
                            ])
                
                split.append([ "Copy Editing", # [0] Stage
                                   data[x][1], # [1] PM
                                   data[x][2], # [2] Campaign
                                   data[x][3], # [3] Prod
                                   data[x][4], # [4] Name, [5] Due Int
                                            7, # [6] Duration
                              (data[x][7]-13), # [7] DueDate
                                   data[x][8], # [8] IsComp
                                   data[x][9], # [9] Sheet
                                  data[x][10], # [10] Quarter
                                  data[x][11], # [11] Type
                            ])

                split.append(["Create Images", # [0] Stage
                                   data[x][1], # [1] PM
                                   data[x][2], # [2] Campaign
                                   data[x][3], # [3] Prod
                                   data[x][4], # [4] Name
                     ((data[x][7]-yearNum)-6), # [5] Due Int
                                           14, # [6] Duration
                               (data[x][7]-6), # [7] DueDate
                                   data[x][8], # [8] IsComp
                                   data[x][9], # [9] Sheet
                                  data[x][10], # [10] Quarter
                                  data[x][11], # [11] Type
                            ])

                split.append(["Eloqua Drafts", # [0] Stage
                                   data[x][1], # [1] PM
                                   data[x][2], # [2] Campaign
                                   data[x][3], # [3] Prod
                                   data[x][4], # [4] Name
                     ((data[x][7]-yearNum)-1), # [5] Due Int
                                            5, # [6] Duration
                               (data[x][7]-1), # [7] DueDate
                                   data[x][8], # [8] IsComp
                                   data[x][9], # [9] Sheet
                                  data[x][10], # [10] Quarter
                                  data[x][11], # [11] Type
                            ])

                split.append([            "!", # [0] Stage
                                   data[x][1], # [1] PM
                                   data[x][2], # [2] Campaign
                                   data[x][3], # [3] Prod
                                   data[x][4], # [4] Name
                         (data[x][7]-yearNum), # [5] Due Int
                                            1, # [6] Duration
                                   data[x][7], # [7] DueDate
                                   data[x][8], # [8] IsComp
                                   data[x][9], # [9] Sheet
                                  data[x][10], # [10] Quarter
                                  data[x][11], # [11] Type
                            ])

        return split                   
                   
    def iscomplete(self, d):
        completed = []

        # Load up the completed deliverables into "c".
        file_location = "Ref\Completed_New.xls"
        workbook = xlrd.open_workbook(file_location) # Open Workbook
        sheet = workbook.sheet_by_index(0) # First Sheet
        c = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in     \
            range(sheet.nrows)]

        # Compare "d" (data) to "c" (completed) and append to "completed"
        for n in range(len(d)):
            for x in range(len(c)):
                if d[n][0] == c[x][0] and d[n][2] == c[x][2]                  \
                and d[n][4] == c[x][4]:
                    d[n][8] = "Y"
            completed.append(d[n])
            
        return completed

class Quarters:
    def __init__(self, obj):
        self.Q1 = self.Quarter1(obj)
        self.Q2 = self.Quarter2(obj)

    # If a row falls in Quarter 1, add it to self.Q1.
    def Quarter1(self, obj):
        data = obj.data
        inrange = []
        for x in range(len(data)):
            if data[x][5] < 92 and (data[x][5]-data[x][6]) > 0:
                inrange.append([ data[x][0], # [0] Stage
                                 data[x][1], # [1] PM
                                 data[x][2], # [2] Campaign
                                 data[x][3], # [3] Prod
                                 data[x][4], # [4] Name
                                 data[x][5], # [5] Due Int
                                 data[x][6], # [6] Duration
                                 data[x][7], # [7] DueDate
                                 data[x][8], # [8] IsComp
                                 data[x][9], # [9] Sheet
                                       "Q1", # [10] Quarter
                                data[x][11], # [11] Type
                            ])
        return inrange

    # If a row falls in Quarter 2, add it to self.Q2.
    def Quarter2(self, obj):
        data = obj.data
        inrange = []
        for x in range(len(data)):
            if data[x][5] < 176 and (data[x][5]-data[x][6]) > 92:
                inrange.append([ data[x][0], # [0] Stage
                                 data[x][1], # [1] PM
                                 data[x][2], # [2] Campaign
                                 data[x][3], # [3] Prod
                                 data[x][4], # [4] Name
                            (data[x][5]-92), # [5] Due Int
                                 data[x][6], # [6] Duration
                                 data[x][7], # [7] DueDate
                                 data[x][8], # [8] IsComp
                                 data[x][9], # [9] Sheet
                                       "Q2", # [10] Quarter
                                data[x][11], # [11] Type
                            ])
        return inrange

class ProductManager:
    def __init__(self, obj, pm_name, prod_name):
        self.n = pm_name
        self.p = prod_name

        """ Quarter 1 """
        self.Q1 = self.mine(obj.Q1)
        self.Q1_xls = self.xlsprep(self.Q1)
        self.Q1_filename = "Final\Deliverables_Q1_" \
            + str(time.strftime("%b%d")) + "_" + str(self.Q1[0][3]) + ".xls"
        self.Q1_gantt = []
        self.gantt(self.Q1, self.Q1_gantt)

        """ Quarter 2 """
        self.Q2 = self.mine(obj.Q2)
        self.Q2_xls = self.xlsprep(self.Q2)
        self.Q2_filename = "Final\Deliverables_Q2_" \
            + str(time.strftime("%b%d")) + "_" + str(self.Q2[0][3]) + ".xls"
        self.Q2_gantt = []
        self.gantt(self.Q2, self.Q2_gantt)
        
        """ Generic Filename """
        self.filename = "Final\Deliverables_" \
            + str(time.strftime("%b%d")) + "_" + str(self.Q1[0][3]) + ".xls"

    def xlsprep(self, data):
        Final = []
        Final.append(["Stage", "Product", "Name", "Due Date", "Complete?"])
        for x in range(len(data)):
            if data[x][1] == self.n:
                Final.append([ data[x][0], # [0] Stage
                               data[x][1], # [1] PM
                               data[x][2], # [2] Campaign
                                   self.p, # [3] Prod
                               data[x][4], # [4] Name
                               data[x][5], # [5] Due Int
                               data[x][6], # [6] Duration
                               data[x][7], # [7] DueDate
                               data[x][8], # [8] IsComp
                               data[x][9], # [9] Sheet
                              data[x][10], # [10] Quarter
                              data[x][11], # [11] Type
                            ])
        return Final

    def mine(self, data):
        Final = []
        for x in range(len(data)):
            if data[x][1] == self.n:
                Final.append([ data[x][0], # [0] Stage
                               data[x][1], # [1] PM
                               data[x][2], # [2] Campaign
                                   self.p, # [3] Prod
                               data[x][4], # [4] Name
                               data[x][5], # [5] Due Int
                               data[x][6], # [6] Duration
                               data[x][7], # [7] DueDate
                               data[x][8], # [8] IsComp
                               data[x][9], # [9] Sheet
                              data[x][10], # [10] Quarter
                              data[x][11], # [11] Type
                            ])
        if len(Final) == 0:
            Final.append(["", "", "", "", "", 1, 1, 1, "", 1, "", ""])
        return Final

    def gantt(self, data, var):
        dat = data # Need to pop items off list, and can't have range change.
        l = len(data)
        ln = l // 24
        m = l % 24
        if m == 0:
            r = range(ln)
        else:
            r = range(ln+1)
        num = 1
        for x in r:
            rng = []
            if len(dat) > 23:
                for y in range(0,23):
                    dat[0][9] = num
                    rng.append(dat[0])
                    dat = dat[1:]
            else:
                for y in range(len(dat)):
                    dat[0][9] = num
                    rng.append(dat[0])
                    dat = dat[1:]
            num += 1
            var.append(rng)

class CampaignRanges:
    def __init__(self, Q_obj):
        self.Q1 = Q_obj.Q1
        self.Q2 = Q_obj.Q2
        
        # Get a list of campaigns.
        self.Q1_clist = self.find_campaigns()
        # List of lists containing campaign deliverables.
        self.Q1_sorted = self.campaign_sort()
        # List of campaigns, with their start and end dates.
        self.Q1_minmax = self.minmax()

        
    def find_campaigns(self):
        Cams = []
        for x in range(len(self.Q1)):
            if self.Q1[x][2] not in Cams:
                Cams.append(self.Q1[x][2])
        return Cams

    def campaign_sort(self):
        Final = []
        for x in range(len(self.Q1_clist)):
            inCampaign = []
            for y in range(len(self.Q1)):
                if self.Q1[y][2] == self.Q1_clist[x]:
                    inCampaign.append(self.Q1[y])
            Final.append(inCampaign)
        return Final

    def minmax(self):
        Minmax = []
        for x in range(len(self.Q1_sorted)):
            large = 0
            small = 99999
            for y in range(len(self.Q1_sorted[x])):
                deliv_end = self.Q1_sorted[x][y][7]
                deliv_start = (self.Q1_sorted[x][y][7] \
                    - self.Q1_sorted[x][y][6])
                if deliv_end > large:
                    large = self.Q1_sorted[x][y][7]
                if  deliv_start < small:
                    small = (self.Q1_sorted[x][y][7] - self.Q1_sorted[x][y][6])
                campaign = [   self.Q1_sorted[x][y][2], # [0] Stage (now camp)
                           self.Q1_sorted[x][y][1], # [1] PM
                           self.Q1_sorted[x][y][2], # [2] Campaign
                           self.Q1_sorted[x][y][3], # [3] Prod
                           self.Q1_sorted[x][y][3], # [4] Name (now prod)
                                   (large - 42003), # [5] Due Int
                                   (large - small), # [6] Duration
                                             large, # [7] DueDate
                           "N", # [8] IsComp
                           self.Q1_sorted[x][y][9], # [9] Sheet
                          "test", # [10] Quarter
                          self.Q1_sorted[x][y][11], # [11] Type
                        ]
            print(campaign)
            Minmax.append(campaign)
        return Minmax

def save(data, str_filename):
    # Create a new workbook object
    newbook = xlwt.Workbook()
    # Add a sheet
    newsheet = newbook.add_sheet('Deliverables')
    # Input data
    for x in range(len(data)):
        for y in range(len(data[x])):
            newsheet.write(x, y, data[x][y])
    # Save file
    newbook.save(str_filename)

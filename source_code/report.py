from timberscale import Timber
import csv
import math

##### REPORT DATA MODULE
class Report(object):
    LOG_RANGE_LIST = [["40+ ft", range(41, 121)], ["31-40 ft", range(31, 41)], ["21-30 ft", range(21, 31)],
                      ["11-20 ft", range(11, 21)], ["1-10 ft", range(1, 11)]]


    def __init__(self, CSV, Stand_to_Examine, Plots, Pref_Log, Min_Log):
        self.csv = CSV
        self.stand = Stand_to_Examine.upper()
        self.plots = Plots
        self.plog = Pref_Log
        self.mlog = Min_Log

        self.species_list = []
        self.summary_conditions = []
        self.summary_logs = []

        self.conditions_dict = {}
        self.logs_dict = {}

        self.report()

    def report(self):
        AVG_HDR, self.species_list = self.get_HDR_Species()

        ## MAIN READ AND TREE (TIMBER CLASS) INITIALIZATION
        with open(self.csv, 'r') as tree_data:
            
            tree_data_reader = csv.reader(tree_data)
            next(tree_data_reader)
            for line in tree_data_reader:
                if line[0] == "":
                    break
                elif str(line[0]).upper() != self.stand:
                    next
                else:
                    SPECIES = str(line[3]).upper()
                    DBH = float(line[4])
                    if line[5] == "":
                        HEIGHT = int(round(AVG_HDR * (DBH/12),0))
                    else:
                        HEIGHT = int(float(line[5]))
                    PLOT_FACTOR = float(line[6])

                    if DBH >= 6.0:
                        tree = Timber(SPECIES, DBH, HEIGHT)

                        merch_dib = tree.merch_dib()
                        if merch_dib < 5:
                            merch_dib = 5
                            
                        single = tree.tree_single(merch_dib, self.plog, self.mlog)
                        tree_per_acre = tree.tree_acre(merch_dib, self.plog, self.mlog, PLOT_FACTOR)
                        log_per_acre = tree.log_acre(merch_dib, self.plog, self.mlog, PLOT_FACTOR)

                        self.summary_conditions.append([single['SPP'][0],
                                                       [tree_per_acre['TPA'], tree_per_acre['BA_AC'], tree_per_acre['RD_AC'],
                                                        single['T_HGT'], single['HDR'], single['VBAR'], tree_per_acre['BF_AC'],
                                                        tree_per_acre['CF_AC']]])


                        self.summary_logs.append(self.get_log_list(single['SPP'][0], log_per_acre))

                    else:
                        tree = Timber(SPECIES, DBH, HEIGHT)

                        self.summary_conditions.append([tree.SPP,
                                                       [tree.get_TPA(PLOT_FACTOR), tree.get_BA_acre(PLOT_FACTOR),
                                                        tree.get_RD_acre(PLOT_FACTOR), tree.HGT, tree.HDR, 0, 0, 0]])
                        


        ## SUMMARY STATISTICS
       
        self.conditions_dict = self.get_conditions_dict()
        self.logs_dict = self.get_logs_dict()

        return


    def get_HDR_Species(self):
        HDR_LIST = []
        SPECIES_LIST = []
        
        with open(self.csv, 'r') as tree_data:
            tree_data_reader = csv.reader(tree_data)
            next(tree_data_reader)

            for line in tree_data_reader:
                if line[0] == "":
                    break
                elif str(line[0]).upper() != self.stand:
                    next
                else:
                    SPP = str(line[3]).upper()
                    if SPP not in SPECIES_LIST:
                        SPECIES_LIST.append(SPP)

                    if line[5] != "":
                        DBH = float(line[4])
                        HEIGHT = float(line[5])
                        HDR_LIST.append(HEIGHT / (DBH / 12))

        AVG_HDR = round(sum(HDR_LIST) / len(HDR_LIST), 2)
        return AVG_HDR, SPECIES_LIST



    def get_log_list(self, Species, Log_Dict):
        master = [Species]
        for key in Log_Dict:
            rng = ""
            for ranges in self.LOG_RANGE_LIST:
                if Log_Dict[key]['L_LGT'] in ranges[1]:
                    rng = ranges[0]

            temp_list = [Log_Dict[key]['L_GRD'][0], rng, Log_Dict[key]['L_CT_AC'],
                         Log_Dict[key]['L_BF_AC'], Log_Dict[key]['L_CF_AC']]

            master.append(temp_list)

        return master



    def get_conditions_dict(self):
        # ORDER OF INITAL SPP LIST - [0SPPCOUNT, 1TPA, 2BA_AC, 3RD_AC, 4T_HGT, 5HDR, 6VBAR, 7BF_AC, 8CF_AC]
        #                             After Pop SPPCOUNT and Add QMD to 2 index
        # ORDER OF FINAL SPP LIST - [0TPA, 1BA_AC, 2QMD, 3RD_AC, 4T_HGT, 5HDR, 6VBAR, 7BF_AC, 8CF_AC]
        master = {}
        totals_temp = [0, 0, 0, 0, 0, 0, 0, 0, 0]
        
        for spp in self.species_list:
            master[spp] = [0, 0, 0, 0, 0, 0, 0, 0, 0]

        for data in self.summary_conditions:
            spp = data[0]
            master[spp][0] += 1
            totals_temp[0] += 1
            for i in range(1, len(data[1]) + 1):
                master[spp][i] += data[1][i - 1]
                totals_temp[i] += data[1][i - 1]

        master["TOTALS"] = totals_temp
        for key in master:
            sums = [1, 2, 3, 7, 8]
            for i in range(1, len(master[key])):
                if i in sums:
                    master[key][i] = master[key][i] / self.plots
                else:
                    master[key][i] = master[key][i] / master[key][0]
            master[key].pop(0)
            master[key].insert(2, math.sqrt((master[key][1] / master[key][0]) / .005454))

        return master


    def get_logs_dict(self):
        log_rng = ["40+ ft", "31-40 ft", "21-30 ft", "11-20 ft", "1-10 ft", 'TGRD']
        master = {}

        # Formatting Species into main keys
        for spp in self.species_list:
            master[spp] = {}

        master['TOTALS'] = {}

        # Formatting Grades and Ranges in correct order, as nested dicts of Species and Totals
        for key in master:
            for grade in Timber.GRADE_NAMES:
                master[key][grade] = {}
                for rng in log_rng:
                    master[key][grade][rng] = [0, 0, 0]
            master[key]['TTL'] = {}
            for rng in log_rng:
                master[key]['TTL'][rng] = [0, 0, 0]

        # Adding data to Master Dict
        for data in self.summary_logs:
            spp = data[0]
                
            for i in range(1, len(data)):
                grade, rng = data[i][0], data[i][1]                

                for j in range(2, len(data[i])):
                    master[spp][grade][rng][j - 2] += (data[i][j] / self.plots)
                    master[spp][grade]['TGRD'][j - 2] += (data[i][j] / self.plots)
                    
                    master[spp]['TTL'][rng][j - 2] += (data[i][j] / self.plots)
                    master[spp]['TTL']['TGRD'][j - 2] += (data[i][j] / self.plots)
                    
                    master['TOTALS'][grade][rng][j - 2] += (data[i][j] / self.plots)
                    master['TOTALS'][grade]['TGRD'][j - 2] += (data[i][j] / self.plots)

                    master['TOTALS']['TTL'][rng][j - 2] += (data[i][j] / self.plots)
                    master['TOTALS']['TTL']['TGRD'][j - 2] += (data[i][j] / self.plots)


        # Removing any Grades that have zero data
        ax_list = []
        for key in master:
            for grade in master[key]:
                count = 0
                for rng in master[key][grade]:
                    count += master[key][grade][rng][0]

                if count == 0:
                    ax_list.append((key, grade))

        for ax in ax_list:
            del master[ax[0]][ax[1]]
                    
        return master













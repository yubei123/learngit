from collections import defaultdict
from typing import DefaultDict
import xlrd
from docxtpl import DocxTemplate, RichText
import sys,pathlib

pp = pathlib.Path(sys.argv[0])
tpl_path = pp.absolute().parent / 'tpl'

infoName = ['xh','kd','report_id','sample_id','hospital','name','gender','age','patient_id','bed_id','tel','hospital_id','department_id','doctor_name','detect_date',\
    'collect_date','jc_date','report_date','proj_type','sample_type','sample_volume','sample_remained','chief_complaint','clinical_diagnosis','pathogen_tip',\
    'drug_list','is_drug_used','wbc','lym','crp','pct','pmn','platelet','culture','identification','scopy']
rgi = { 'acridine dye': "二氯基吖啶", "aminocoumarin antibiotic": "氨基香豆素类", "aminoglycoside antibiotic": "氨基糖苷类", "antibacterial free fatty acids": "FFA抗菌素游离脂肪酸", \
    "benzalkonium chloride": "苯扎氯铵", "bicyclomycin": "双环霉素", "carbapenem": "碳青霉烯类", "cephalosporin": "头孢菌素", "cephamycin": "头霉素类", \
    "cycloserine": "环丝氨酸", "diaminopyrimidine antibiotic": "促生长类", "diarylquinoline antibiotic": "二芳基喹啉", "elfamycin antibiotic": "elfamycin类", \
    "ethionamide": "乙硫异烟胺", "fluoroquinolone antibiotic": "氟喹诺酮类", "fosfomycin": "磷霉素", "fosmidomycin": "膦胺霉素", "fusidic acid": "夫西地酸", \
    "glycopeptide antibiotic": "糖肽类", "glycylcycline": "甘氨酰胺四环素", "isoniazid": "异烟肼", "lincosamide antibiotic": "林可霉素类", "macrocyclic antibiotic": "大环类", \
    'macrolide antibiotic': "大环内酯类", "monobactam": "单内酰环类", "mupirocin": "莫匹罗星", "nitrofuran antibiotic": "硝基呋喃", "nitroimidazole antibiotic": "硝基咪唑", \
    "nucleoside antibiotic": "核苷类抗生素", "nybomycin": "尼博霉素", "organoarsenic antibiotic": "有机砷", "oxazolidinone antibiotic": "噁唑烷酮抗生素", "pactamycin": "约霉素", \
    "para-aminosalicylic acid": "对氨水杨酸", "penam": "青霉烷类", "penem": "青霉烯类", "peptide antibiotic": "肽抗生素类", "phenicol antibiotic": "苯丙醇", \
    "pleuromutilin antibiotic": "截短侧耳素类", "polyamine antibiotic": "多胺", "prothionamide": "丙硫异烟胺", "pyrazinamide": "吡嗪酰胺", "rhodamine": "罗丹明", \
    "rifamycin antibiotic": "利福霉素", "streptogramin antibiotic": "链阳霉素类", "sulfonamide antibiotic": "磺胺类", "sulfone antibiotic": "砜类抗生素", \
    "tetracycline antibiotic": "四环素", "triclosan": "三氯生类", "penicillin": "青霉素类", "chloramphenic": "氯霉素类", "Minocycline": "米诺环素"}
med = {'antibiotic efflux':'抗生素外排', 'antibiotic target alteration':'抗生素靶点改变', 'antibiotic inactivation':'抗生素灭活', 'reduced permeability to antibiotic':'抗生素渗透性降低',\
    'antibiotic target protection':'抗生素靶点保护','antibiotic target replacement':'抗生素靶点置换'}

possample = []
negsample = []
rgisample = []
hysample = []
mzsample = []
boaosample = []
report_date = {}
report_sample = []
the_type = defaultdict(str)
sample_type = defaultdict(str)

def getSampleInfo(inputfile,rgidir):
    sn = 1
    num = 6
    samplesn = defaultdict(int)
    samplenum = defaultdict(int)
    book = xlrd.open_workbook(inputfile)
    allsample = []
    sample = defaultdict(dict)

    ##读取BASIC sheet
    infosheet = book.sheet_by_index(0)
    for i in range(1, infosheet.nrows):
        sample_id = infosheet.row(i)[3].value.strip()
        allsample.append(sample_id)
        report_date[sample_id] = infosheet.row(i)[17].value.strip()
        the_type[sample_id] = infosheet.row(i)[18].value.strip()
        sample[sample_id].update({ 'the_type':the_type[sample_id] })
        sample_type[sample_id] = infosheet.row(i)[19].value.strip()
        samplesn[sample_id] = sn
        samplenum[sample_id] = num
        for m, n in zip(infosheet.row(i), infoName):
            value = str(m.value).strip()
            if not value:
                value = '-'
            sample[sample_id].update({ n: value })
    
    ##读取模版信息
    tplsheet = book.sheet_by_index(1)
    for i in range(1, tplsheet.nrows):
        tpl = [str(j.value).strip() for j in tplsheet.row(i)]
        sample[tpl[0]].update({ 'tpl':tpl[-1] })
        if tpl[1] == '阳性':
            possample.append(tpl[0])
        else:
            negsample.append(tpl[0])
        if tpl[0] in allsample:
            report_sample.append(tpl[0])
            if tpl[-1].find('hy') > -1:
                print(f'华银报告：{tpl[0]}')                
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                hysample.append(tpl[0])
            elif tpl[-1].find('mz') > -1:
                print(f'明志报告：{tpl[0]}')                
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                mzsample.append(tpl[0])
            elif tpl[-1].find('boao') > -1:
                print(f'博奥报告：{tpl[0]}')                
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                boaosample.append(tpl[0])
            else:
                print(f'aja/zju/nj 报告：{tpl[0]}')
                sample[tpl[0]].update({ 'project_type':tpl[2] })
                rgisample.append(tpl[0])
                if tpl[-1].find('zju') > -1:
                    if tpl[-1].find('positive2') > -1 or tpl[-1].find('negative2') > -1:
                        sample[tpl[0]].update({'shuiyin':'免费测试','the_detect':'','the_signature':'','the_report_date':'','date':'','beizhu':''})
                    else:
                        if str(tpl[-1]).find('pay') > -1:
                            sample[tpl[0]].update({'shuiyin':'','the_detect':'检测者：','the_signature':'审核签字：','the_report_date':'报告日期：','date':report_date[tpl[0]],'beizhu':'备注：此报告仅对本次送检样本负责！结果仅供医生参考。\n若对检测结果有疑问，请于收到报告后7个工作日内与我们联系，谢谢合作！'})
                        else:
                            sample[tpl[0]].update({'shuiyin':'免费测试','the_detect':'检测者：','the_signature':'审核签字：','the_report_date':'报告日期：','date':report_date[tpl[0]],'beizhu':'备注：此报告仅对本次送检样本负责！结果仅供医生参考。\n若对检测结果有疑问，请于收到报告后7个工作日内与我们联系，谢谢合作！'})
                elif tpl[-1].find('aja') > -1:
                    if tpl[-1].find('pay') > -1:
                        sample[tpl[0]].update({'shuiyin':''})
                    else:
                        sample[tpl[0]].update({'shuiyin':'免费测试'})

    ##读取阳性病原体信息
    possheet = book.sheet_by_index(2)
    highBacteria,highVirus,highFungi,highParasite,highSpecial = defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list)
    lowBacteria,lowVirus,lowFungi,lowParasite,lowSpecial = defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list)
    virusList,bacteriaList,fungiList,parasiteList,specialList = defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list),defaultdict(list)
    description,papers = defaultdict(list),defaultdict(list)
    highList,lowList = defaultdict(list),defaultdict(list)
    fungi_parasiteList,bacteria_specialList = defaultdict(list),defaultdict(list)
    library = defaultdict(str)

    sampleposReads = defaultdict(lambda: defaultdict(int))
    sampleposInfos = defaultdict(lambda: defaultdict(list))
    allmicro = defaultdict(lambda: defaultdict(str))
        
    for i in range(1,possheet.nrows):
        pos = [str(j.value).strip() for j in possheet.row(i)]
        if pos[0] in possample:
            sampleposReads[pos[0]][pos[8]] = int(float(pos[4]))
            sampleposInfos[pos[0]][pos[8]] = pos
    
    area = defaultdict(lambda: defaultdict(list))
    znen = defaultdict(lambda: defaultdict(list))
    species_type = defaultdict(lambda: defaultdict(list))
    for i in sampleposReads:
        for s, v in sorted(sampleposReads[i].items(), key=lambda x: x[1], reverse=True):
            pos = sampleposInfos[i][s]   
            species_type[pos[0]][pos[10]] = pos[2]         
            speRT2 = RichText()
            speRT2.add(str(pos[8]) + '\n')
            speRT2.add(str(pos[3]),italic=True)
            genusRT = RichText()
            genusRT.add(str(pos[9]) + '\n')
            genusRT.add(str(pos[10]), italic=True)
            abu_raw = float(pos[6])
            abu_clean = str(float('%.3f' % float(abu_raw))) if abu_raw > 0.001 else str('&lt;' + '0.001')
            e_sp = {'type': str(pos[-3]),
                    'species': speRT2,
                    'scount': format(int(float(pos[4])),','),
                    'abundance': str(abu_clean) + str('%'),
                    'focus': str(pos[7])}
            znen[pos[0]][pos[10]] = {'genus':genusRT, 'gcount': format(int(float(pos[5])), ',')}
            if pos[10] in znen[pos[0]]:
                area[pos[0]][pos[10]].append(e_sp) 
            else:
                area[pos[0]][pos[10]] = [e_sp]

    for i in area:
        for s, v in area[i].items():
            znen[i][s]['area'] = area[i][s]
            if species_type[i][s] == 'fungi':  
                fungiList[i].append(znen[i][s])
            elif species_type[i][s] == 'bacteria':
                bacteriaList[i].append(znen[i][s])
            elif species_type[i][s] == 'virus':
                virusList[i].append(znen[i][s])
            elif species_type[i][s] == 'parasite':
                parasiteList[i].append(znen[i][s])
            elif species_type[i][s] == 'special':
                specialList[i].append(znen[i][s])
            else:
                print(f'{i}物种类型单词写错了！')
            fungi_parasiteList[i] = fungiList[i] + parasiteList[i]
            bacteria_specialList[i] = bacteriaList[i] + specialList[i]
    
    for i in sampleposReads:
        for s, v in sorted(sampleposReads[i].items(), key=lambda x: x[1], reverse=True):
            pos = sampleposInfos[i][s]
            library[pos[0]] = pos[1]
            microRT = RichText()
            microRT.add(str(pos[8]) + '\n')
            microRT.add(str(pos[3]),italic=True)
            allmicro[pos[0]][pos[3]] = microRT
            speRT = RichText()
            speRT.add(str(pos[8]) + ' ( ')
            speRT.add(str(pos[3]),italic=True)
            speRT.add(' )')
            descRT = RichText()
            paperRT = RichText()
            if pos[0] in rgisample:
                descRT.add(str(samplesn[pos[0]]) + '. ')
                descRT.add(pos[8], color='#1BB8CE',bold=True)
                descRT.add(' ( ')
                descRT.add(pos[3],italic=True)
                descRT.add(' )')
                descRT.add('[' + str(samplenum[pos[0]]) + ']',style='sup')
                paperRT.add('[' + str(samplenum[pos[0]]) + '] ',style='paper')
                if pos[-2] != '' and pos[-1] != '':
                    descRT.add(' : ' + f'{pos[-2]}')
                    paperRT.add(str(pos[-1]),style='paper')
                elif pos[-2] != '' and pos[-1] == '':
                    descRT.add(' : ' + f'{pos[-2]}')
                    paperRT.add('NA',style='paper')
                else:
                    descRT.add(' : ' + 'NA')
                    paperRT.add('NA',style='paper')
                papers[pos[0]].append(paperRT)
                description[pos[0]].append(descRT)
            elif pos[0] in hysample:
                descRT.add(str(samplesn[pos[0]]) + '. ')
                descRT.add(pos[8], color='#0079CA',bold=True)
                descRT.add(' ( ')
                descRT.add(pos[3],italic=True)
                descRT.add(' )')
                descRT.add('[' + str(samplenum[pos[0]]) + ']',style='sup')
                paperRT.add('[' + str(samplenum[pos[0]]) + '] ',style='paper')
                if pos[-2] != '' and pos[-1] != '':                    
                    descRT.add(' : ' + f'{pos[-2]}')
                    paperRT.add(str(pos[-1]),style='paper')
                elif pos[-2] != '' and pos[-1] == '':
                    descRT.add(' : ' + f'{pos[-2]}')
                    paperRT.add('NA',style='paper')
                else:
                    descRT.add(' : ' + 'NA')
                    paperRT.add('NA',style='paper')
                papers[pos[0]].append(paperRT)
                description[pos[0]].append(descRT)
            elif pos[0] in boaosample:
                descRT.add(pos[8],bold=True)
                descRT.add('(')
                descRT.add(pos[3],italic=True,bold=True)
                descRT.add(')')
                if pos[-2] != '' and pos[-1] != '':
                    descRT.add(' : ' + f'{pos[-2]}')
                elif pos[-2] != '' and pos[-1] == '':
                    descRT.add(' : ' + f'{pos[-2]}')
                else:
                    descRT.add(' : ' + 'NA')
                description[pos[0]].append(descRT)
            else:
                descRT.add(str(samplesn[pos[0]]) + '. ')
                descRT.add(pos[8],bold=True)
                descRT.add(' ( ')
                descRT.add(pos[3],italic=True)
                descRT.add(' )')
                descRT.add('[' + str(samplenum[pos[0]]) + ']',style='sup')
                paperRT.add('[' + str(samplenum[pos[0]]) + '] ',style='paper')
                if pos[-2] != '' and pos[-1] != '':
                    descRT.add(' : ' + f'{pos[-2]}')
                    paperRT.add(str(pos[-1]),style='paper')
                elif pos[-2] != '' and pos[-1] == '':
                    descRT.add(' : ' + f'{pos[-2]}')
                    paperRT.add('NA',style='paper')
                else:
                    descRT.add(' : ' + 'NA')
                    paperRT.add('NA',style='paper')
                papers[pos[0]].append(paperRT)
                description[pos[0]].append(descRT)
            samplesn[pos[0]] += 1
            samplenum[pos[0]] += 1            
            if pos[2] == 'bacteria':
                highBacteria[pos[0]].append({'bacteria':'细菌','species':speRT}) if pos[7] == '高' else lowBacteria[pos[0]].append({'bacteria':'细菌','species':speRT})
            elif pos[2] == 'virus':
                highVirus[pos[0]].append({'virus':'病毒','species':speRT}) if pos[7] == '高' else lowVirus[pos[0]].append({'virus':'病毒','species':speRT})
            elif pos[2] == 'fungi':
                highFungi[pos[0]].append({'fungi':'真菌','species':speRT}) if pos[7] == '高' else lowFungi[pos[0]].append({'fungi':'真菌','species':speRT})
            elif pos[2] == 'parasite':
                highParasite[pos[0]].append({'parasite':'寄生虫','species':speRT}) if pos[7] == '高' else lowParasite[pos[0]].append({'parasite':'寄生虫','species':speRT})
            elif pos[2] == 'special':
                highSpecial[pos[0]].append({'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':speRT}) if pos[7] == '高' else lowSpecial[pos[0]].append({'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':speRT})
            else:
                print(f'{pos[2]}物种类型单词写错了！')
            highList[pos[0]] =  highBacteria[pos[0]] + highVirus[pos[0]] + highFungi[pos[0]] + highParasite[pos[0]] + highSpecial[pos[0]]
            lowList[pos[0]] = lowBacteria[pos[0]] + lowVirus[pos[0]] + lowFungi[pos[0]] + lowParasite[pos[0]] + lowSpecial[pos[0]]

    ##读取阴性疑似病原体和背景微生物信息
    negsheet = book.sheet_by_index(3)
    backlist = defaultdict(list)
    sampleneg_by_Reads = defaultdict(lambda: defaultdict(int))
    sampleneg_by_Infos = defaultdict(lambda: defaultdict(list))
    sampleneg_bj_Reads = defaultdict(lambda: defaultdict(int))
    sampleneg_bj_Infos = defaultdict(lambda: defaultdict(list))
    
    for i in range(1,negsheet.nrows):
        neg = [str(j.value).strip() for j in negsheet.row(i)]
        if neg[0] in rgisample or neg[0] in hysample or neg[0] in mzsample:
            if str(neg[5]) == '疑似病原体':
                sampleneg_by_Reads[neg[0]][neg[2]] = int(float(neg[4]))
                sampleneg_by_Infos[neg[0]][neg[2]] = neg
            else:
                sampleneg_bj_Reads[neg[0]][neg[2]] = int(float(neg[4]))
                sampleneg_bj_Infos[neg[0]][neg[2]] = neg
    
    for i in sampleneg_by_Reads:
        for s, v in sorted(sampleneg_by_Reads[i].items(), key=lambda x: x[1], reverse=True):
            neg = sampleneg_by_Infos[i][s]
            speRT = RichText()
            speRT.add(str(neg[3]) + '\n( ')
            speRT.add(str(neg[2]),italic=True)
            speRT.add(' )')
            descRT = RichText()
            paperRT = RichText()
            if neg[0] in rgisample:
                descRT.add(str(samplesn[neg[0]]) + '. ')
                descRT.add(neg[3], color='#1BB8CE',bold=True)
                descRT.add(' ( ')
                descRT.add(neg[2],italic=True)
                descRT.add(' )')
                descRT.add('[' + str(samplenum[neg[0]]) + ']',style='sup')
                paperRT.add('[' + str(samplenum[neg[0]]) + '] ',style='paper')
                if neg[-2] != '' and neg[-1] != '':
                    descRT.add(' : ' + f'{neg[-2]}')
                    paperRT.add(str(neg[-1]),style='paper')
                elif neg[-2] != '' and neg[-1] == '':
                    descRT.add(' : ' + f'{neg[-2]}')
                    paperRT.add('NA',style='paper')
                else:
                    descRT.add(' : ' + 'NA')
                    paperRT.add('NA',style='paper')
                papers[neg[0]].append(paperRT)
                description[neg[0]].append(descRT)
            elif neg[0] in hysample:
                descRT.add(str(samplesn[neg[0]]) + '. ')
                descRT.add(neg[3], color='#0079CA',bold=True)
                descRT.add(' ( ')
                descRT.add(neg[2],italic=True)
                descRT.add(' )')
                descRT.add('[' + str(samplenum[neg[0]]) + ']',style='sup')
                paperRT.add('[' + str(samplenum[neg[0]]) + '] ',style='paper')
                if neg[-2] != '' and neg[-1] != '':
                    descRT.add(' : ' + f'{neg[-2]}')
                    paperRT.add(str(neg[-1]),style='paper')
                elif neg[-2] != '' and neg[-1] == '':
                    descRT.add(' : ' + f'{neg[-2]}')
                    paperRT.add('NA',style='paper')
                else:
                    descRT.add(' : ' + 'NA')
                    paperRT.add('NA',style='paper')
                papers[neg[0]].append(paperRT)
                description[neg[0]].append(descRT)
            elif neg[0] in mzsample:
                descRT.add(str(samplesn[neg[0]]) + '. ')
                descRT.add(neg[3],bold=True)
                descRT.add(' ( ')
                descRT.add(neg[2],italic=True)
                descRT.add(' )')
                descRT.add('[' + str(samplenum[neg[0]]) + ']',style='sup')
                paperRT.add('[' + str(samplenum[neg[0]]) + '] ',style='paper')
                if neg[-2] != '' and neg[-1] != '':
                    descRT.add(' : ' + f'{neg[-2]}')
                    paperRT.add(str(neg[-1]),style='paper')
                elif neg[-2] != '' and neg[-1] == '':
                    descRT.add(' : ' + f'{neg[-2]}')
                    paperRT.add('NA',style='paper')
                else:
                    descRT.add(' : ' + 'NA')
                    paperRT.add('NA',style='paper')
                papers[neg[0]].append(paperRT)
                description[neg[0]].append(descRT)
            samplesn[neg[0]] += 1
            samplenum[neg[0]] += 1
            backlist[neg[0]].append({'type':neg[1],'microbe':speRT,'count':f'{int(float(neg[4])):,}','note':neg[5]})
    
    for i in sampleneg_bj_Reads:
        for s, v in sorted(sampleneg_bj_Reads[i].items(), key=lambda x: x[1], reverse=True):
            neg = sampleneg_bj_Infos[i][s]
            speRT = RichText()
            speRT.add(str(neg[3]) + '\n( ')
            speRT.add(str(neg[2]),italic=True)
            speRT.add(' )')
            descRT = RichText()
            paperRT = RichText()
            if neg[-1] != '':
                descRT.add(str(samplesn[neg[0]]) + '. ')
                if neg[0] in rgisample:
                    descRT.add(neg[3], color='#1BB8CE',bold=True)
                elif neg[0] in hysample:
                    descRT.add(neg[3], color='#0079CA',bold=True)
                else:
                    descRT.add(neg[3], bold=True)
                descRT.add(' ( ')
                descRT.add(neg[2],italic=True)
                descRT.add(' )')
                descRT.add('[' + str(samplenum[neg[0]]) + ']',style='sup')
                descRT.add(' : ' + f'{neg[-2]}')
                description[neg[0]].append(descRT)
                paperRT.add('[' + str(samplenum[neg[0]]) + '] ',style='paper')
                paperRT.add(str(neg[-1]),style='paper')   
                papers[neg[0]].append(paperRT)  
            samplesn[neg[0]] += 1
            samplenum[neg[0]] += 1
            backlist[neg[0]].append({'type':neg[1],'microbe':speRT,'count':f'{int(float(neg[4])):,}','note':neg[5]})
            
    ##模版内容添加
    amr_summary = ''
    for i in sample:
        if i in possample:
            if i in rgisample or i in hysample:
                sample[i].update({ 'report_type':'检出以下疑似病原体'})
                sample[i].update({ 'highBacteria':highBacteria[i] }) if highBacteria[i] else sample[i].update({ 'highBacteria':[{'bacteria':'细菌','species':RichText('-')}]})
                sample[i].update({ 'lowBacteria':lowBacteria[i] }) if lowBacteria[i] else sample[i].update({ 'lowBacteria':[{'bacteria':'细菌','species':RichText('-')}] })
                sample[i].update({ 'highVirus':highVirus[i] }) if highVirus[i] else sample[i].update({ 'highVirus':[{'virus':'病毒','species':RichText('-')}] })
                sample[i].update({ 'lowVirus':lowVirus[i] }) if lowVirus[i] else sample[i].update({ 'lowVirus':[{'virus':'病毒','species':RichText('-')}] })
                sample[i].update({ 'highFungi':highFungi[i] }) if highFungi[i] else sample[i].update({ 'highFungi':[{'fungi':'真菌','species':RichText('-')}] })
                sample[i].update({ 'lowFungi':lowFungi[i] }) if lowFungi[i] else sample[i].update({ 'lowFungi':[{'fungi':'真菌','species':RichText('-')}] })
                sample[i].update({ 'highParasite':highParasite[i] }) if highParasite[i] else sample[i].update({ 'highParasite':[{'parasite':'寄生虫','species':RichText('-')}] })
                sample[i].update({ 'lowParasite':lowParasite[i] }) if lowParasite[i] else sample[i].update({ 'lowParasite':[{'parasite':'寄生虫','species':RichText('-')}] })
                sample[i].update({ 'highSpecial':highSpecial[i] }) if highSpecial[i] else sample[i].update({ 'highSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('-')}] })
                sample[i].update({ 'lowSpecial':lowSpecial[i] }) if lowSpecial[i] else sample[i].update({ 'lowSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('-')}] })
                sample[i].update({ 'bacteriaList':bacteriaList[i] }) if bacteriaList[i] else sample[i].update({ 'bacteriaList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'virusList':virusList[i] }) if virusList[i] else sample[i].update({ 'virusList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'fungiList':fungiList[i] }) if fungiList[i] else sample[i].update({ 'fungiList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'parasiteList':parasiteList[i] }) if parasiteList[i] else sample[i].update({ 'parasiteList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'specialList':specialList[i] }) if specialList[i] else sample[i].update({ 'specialList':[{'genus':RichText('-'),'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[{'type':'-', 'microbe':RichText('-'), 'count':'-', 'note':'-'}] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
                with open(f'{rgidir}/{library[i]}.gene_mapping_data.txt') as rgifile:
                    fh = rgifile.readlines()
                    if len(fh) == 1:
                        amr_summary = '通过分析，未检出耐药基因。'
                        sample[i].update({'amr_summary':amr_summary})
                        sample[i].update({ 'amr':[{'species':RichText('-'), 'area':[{'mechanisms':'-', 'gene':RichText('-'), 'count':'-', 'coverage':'-', 'drug':'-'}]}] })
                    else:
                        amr = defaultdict(lambda: defaultdict(list))
                        amr_area = defaultdict(lambda: defaultdict(list))
                        flag = 0
                        for j in fh[1:]:
                            e = j.strip().split('\t')
                            if allmicro[i][e[-1]]:
                                flag = 1
                                amr_summary = '通过分析，发现患者可能对以下抗生素耐药。'
                                sample[i].update({'amr_summary':amr_summary})
                                mechanisms = ';'.join([med[x] for x in e[-2].split('; ')])    
                                species = allmicro[i][str(e[-1])]
                                drugs = e[4].split('; ')
                                rgis = '; '.join([rgi[x] for x in drugs])
                                genename = RichText(e[1], italic=True)
                                coverage = str(float('%.1f' % float(e[3]))) + str('%')
                                amr[i][e[-1]] = { 'species':species }
                                if len(drugs) <= 3:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':rgis }
                                else:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':'多重耐药' }
                                if e[-1] in amr[i]:
                                    amr_area[i][e[-1]].append(e_sp)
                                else:
                                    amr_area[i][e[-1]] = [e_sp]
                        if flag == 1:
                            b = []
                            for k,v in amr_area[i].items():
                                amr[i][k]['area'] = amr_area[i][k]
                                b.append(amr[i][k])
                            sample[i].update({ 'amr':b })                    
                        elif flag == 0:
                            amr_summary = '通过分析，未检出耐药基因。'
                            sample[i].update({'amr_summary':amr_summary})
                            sample[i].update({ 'amr':[{'species':RichText('-'), 'area':[{'mechanisms':'-', 'gene':RichText('-'), 'count':'-', 'coverage':'-', 'drug':'-'}]}] })
            elif i in boaosample:
                sample[i].update({ 'report_type':'检出以上疑似病原体'})
                sample[i].update({ 'highList':highList[i] }) if highList[i] else sample[i].update({ 'highList':[{'species':RichText('-')}] })
                sample[i].update({ 'lowList':lowList[i] }) if lowList[i] else sample[i].update({ 'lowList':[{'species':RichText('-')}] })
                sample[i].update({ 'bacteria_specialList':bacteria_specialList[i] }) if bacteria_specialList[i] else sample[i].update({ 'bacteria_specialList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'virusList':virusList[i] }) if virusList[i] else sample[i].update({ 'virusList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'fungi_parasiteList':fungi_parasiteList[i] }) if fungi_parasiteList[i] else sample[i].update({ 'fungi_parasiteList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[{'type':'-', 'microbe':RichText('-'), 'count':'-', 'note':'-'}] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'amr':[{'species':RichText('-'), 'area':[{'mechanisms':'-', 'gene':RichText('-'), 'count':'-', 'coverage':'-', 'drug':'-'}]}] })
                with open(f'{rgidir}/{library[i]}.gene_mapping_data.txt') as rgifile:
                    fh = rgifile.readlines()
                    if len(fh) == 1:
                        sample[i].update({ 'amr':[{'species':RichText('-'), 'area':[{'mechanisms':'-', 'gene':RichText('-'), 'count':'-', 'coverage':'-', 'drug':'-'}]}] })
                    else:
                        amr = defaultdict(lambda: defaultdict(list))
                        amr_area = defaultdict(lambda: defaultdict(list))
                        flag = 0
                        for j in fh[1:]:
                            e = j.strip().split('\t')
                            if allmicro[i][str(e[-1])]:
                                flag = 1
                                amr_summary = '通过分析，发现患者可能对以下抗生素耐药。'
                                sample[i].update({'amr_summary':amr_summary})
                                mechanisms = ';'.join([med[x] for x in e[-2].split('; ')])    
                                species = allmicro[i][str(e[-1])]
                                drugs = e[4].split('; ')
                                rgis = '; '.join([rgi[x] for x in drugs])
                                genename = RichText(e[1], italic=True)
                                coverage = str(float('%.1f' % float(e[3]))) + str('%')
                                amr[i][e[-1]] = { 'species':species }
                                if len(drugs) <= 3:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':rgis }
                                else:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':'多重耐药' }
                                if e[-1] in amr[i]:
                                    amr_area[i][e[-1]].append(e_sp)
                                else:
                                    amr_area[i][e[-1]] = [e_sp]
                        if flag == 1:
                            b = []
                            for k,v in amr_area[i].items():
                                amr[i][k]['area'] = amr_area[i][k]
                                b.append(amr[i][k])
                            sample[i].update({ 'amr':b })                    
                        elif flag == 0:
                            sample[i].update({ 'amr':[{'species':RichText('-'), 'area':[{'mechanisms':'-', 'gene':RichText('-'), 'count':'-', 'coverage':'-', 'drug':'-'}]}] })
            else:
                sample[i].update({ 'report_type':'检出以下疑似病原体'})
                sample[i].update({ 'highBacteria':highBacteria[i] }) if highBacteria[i] else sample[i].update({ 'highBacteria':[{'bacteria':'细菌','species':RichText('未检出')}]})
                sample[i].update({ 'lowBacteria':lowBacteria[i] }) if lowBacteria[i] else sample[i].update({ 'lowBacteria':[{'bacteria':'细菌','species':RichText('未检出')}] })
                sample[i].update({ 'highVirus':highVirus[i] }) if highVirus[i] else sample[i].update({ 'highVirus':[{'virus':'病毒','species':RichText('未检出')}] })
                sample[i].update({ 'lowVirus':lowVirus[i] }) if lowVirus[i] else sample[i].update({ 'lowVirus':[{'virus':'病毒','species':RichText('未检出')}] })
                sample[i].update({ 'highFungi':highFungi[i] }) if highFungi[i] else sample[i].update({ 'highFungi':[{'fungi':'真菌','species':RichText('未检出')}] })
                sample[i].update({ 'lowFungi':lowFungi[i] }) if lowFungi[i] else sample[i].update({ 'lowFungi':[{'fungi':'真菌','species':RichText('未检出')}] })
                sample[i].update({ 'highParasite':highParasite[i] }) if highParasite[i] else sample[i].update({ 'highParasite':[{'parasite':'寄生虫','species':RichText('未检出')}] })
                sample[i].update({ 'lowParasite':lowParasite[i] }) if lowParasite[i] else sample[i].update({ 'lowParasite':[{'parasite':'寄生虫','species':RichText('未检出')}] })
                sample[i].update({ 'highSpecial':highSpecial[i] }) if highSpecial[i] else sample[i].update({ 'highSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('未检出')}] })
                sample[i].update({ 'lowSpecial':lowSpecial[i] }) if lowSpecial[i] else sample[i].update({ 'lowSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('未检出')}] })
                sample[i].update({ 'bacteriaList':bacteriaList[i] })
                sample[i].update({ 'virusList':virusList[i] })
                sample[i].update({ 'fungiList':fungiList[i] })
                sample[i].update({ 'parasiteList':parasiteList[i] })
                sample[i].update({ 'specialList':specialList[i] })
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[{'type':'-', 'microbe':RichText('-'), 'count':'-', 'note':'-'}] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
                with open(f'{rgidir}/{library[i]}.gene_mapping_data.txt') as rgifile:
                    fh = rgifile.readlines()
                    if len(fh) == 1:
                        amr = []
                    else:
                        amr = defaultdict(lambda: defaultdict(list))
                        amr_area = defaultdict(lambda: defaultdict(list))
                        flag = 0
                        for j in fh[1:]:
                            e = j.strip().split('\t')
                            if allmicro[i][str(e[-1])]:
                                flag = 1
                                amr_summary = '通过分析，发现患者可能对以下抗生素耐药。'
                                sample[i].update({'amr_summary':amr_summary})
                                mechanisms = ';'.join([med[x] for x in e[-2].split('; ')])    
                                species = allmicro[i][str(e[-1])]
                                drugs = e[4].split('; ')
                                rgis = '; '.join([rgi[x] for x in drugs])
                                genename = RichText(e[1], italic=True)
                                coverage = str(float('%.1f' % float(e[3]))) + str('%')
                                amr[i][e[-1]] = { 'species':species }
                                if len(drugs) <= 3:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':rgis }
                                else:
                                    e_sp = { 'mechanisms':mechanisms, 'gene': genename, 'count': e[2].replace('.00', ''), 'coverage':coverage, 'drug':'多重耐药' }
                                if e[-1] in amr[i]:
                                    amr_area[i][e[-1]].append(e_sp)
                                else:
                                    amr_area[i][e[-1]] = [e_sp]
                        if flag == 1:
                            b = []
                            for k,v in amr_area[i].items():
                                amr[i][k]['area'] = amr_area[i][k]
                                b.append(amr[i][k])
                            sample[i].update({ 'amr':b })                    
                        elif flag == 0:
                            amr = []                            
        elif i in negsample:
            if i in rgisample or i in hysample:
                sample[i].update({ 'report_type':'未检出明确的病原微生物'})
                sample[i].update({ 'highBacteria':[{'bacteria':'细菌','species':RichText('-')}]})
                sample[i].update({ 'lowBacteria':[{'bacteria':'细菌','species':RichText('-')}] })
                sample[i].update({ 'highVirus':[{'virus':'病毒','species':RichText('-')}] })
                sample[i].update({ 'lowVirus':[{'virus':'病毒','species':RichText('-')}] })
                sample[i].update({ 'highFungi':[{'fungi':'真菌','species':RichText('-')}] })
                sample[i].update({ 'lowFungi':[{'fungi':'真菌','species':RichText('-')}] })
                sample[i].update({ 'highParasite':[{'parasite':'寄生虫','species':RichText('-')}] })
                sample[i].update({ 'lowParasite':[{'parasite':'寄生虫','species':RichText('-')}] })
                sample[i].update({ 'highSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('-')}] })
                sample[i].update({ 'lowSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('-')}] })
                sample[i].update({ 'bacteriaList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'virusList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'fungiList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'parasiteList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'specialList':[{'genus':RichText('-'),'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[{'type':'-', 'microbe':RichText('-'), 'count':'-', 'note':'-'}] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
                sample[i].update({ 'amr_summary':'通过分析，未检出耐药基因。' })
                sample[i].update({ 'amr':[{'species':RichText('-'), 'area':[{'mechanisms':'-', 'gene':RichText('-'), 'count':'-', 'coverage':'-', 'drug':'-'}]}] })
            elif i in boaosample:
                sample[i].update({ 'report_type':'未检出明确的病原微生物'})
                sample[i].update({ 'highList':[{'species':RichText('-')}] })
                sample[i].update({ 'lowList':[{'species':RichText('-')}] })
                sample[i].update({ 'bacteria_specialList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'virusList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'type':'-', 'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'fungi_parasiteList':[{'genus':RichText('-'), 'gcount':'-', 'area':[{'species':RichText('-'), 'scount':'-', 'abundance':'-', 'focus':'-'}]}]})
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[{'type':'-', 'microbe':RichText('-'), 'count':'-', 'note':'-'}] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'amr':[{'species':RichText('-'), 'area':[{'mechanisms':'-', 'gene':RichText('-'), 'count':'-', 'coverage':'-', 'drug':'-'}]}] })
            else:
                amr = []
                sample[i].update({ 'report_type':'未检出明确的病原微生物'})
                sample[i].update({ 'highBacteria':[{'bacteria':'细菌','species':RichText('未检出')}]})
                sample[i].update({ 'lowBacteria':[{'bacteria':'细菌','species':RichText('未检出')}] })
                sample[i].update({ 'highVirus':[{'virus':'病毒','species':RichText('未检出')}] })
                sample[i].update({ 'lowVirus':[{'virus':'病毒','species':RichText('未检出')}] })
                sample[i].update({ 'highFungi':[{'fungi':'真菌','species':RichText('未检出')}] })
                sample[i].update({ 'lowFungi':[{'fungi':'真菌','species':RichText('未检出')}] })
                sample[i].update({ 'highParasite':[{'parasite':'寄生虫','species':RichText('未检出')}] })
                sample[i].update({ 'lowParasite':[{'parasite':'寄生虫','species':RichText('未检出')}] })
                sample[i].update({ 'highSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('未检出')}] })
                sample[i].update({ 'lowSpecial':[{'special':'特殊病原体（包括分枝杆菌、支原体/衣原体等）','species':RichText('未检出')}] })
                sample[i].update({ 'backlist':backlist[i] }) if backlist[i] else sample[i].update({ 'backlist':[{'type':'-', 'microbe':RichText('-'), 'count':'-', 'note':'-'}] })
                sample[i].update({ 'descriptions':description[i] }) if description[i] else sample[i].update({ 'descriptions':'-' })
                sample[i].update({ 'papers':papers[i] }) if papers[i] else sample[i].update({ 'papers':'-' })
    
    ##读取数据量信息
    runstatsheet = book.sheet_by_index(5)
    total_reads,human_reads,micro_reads,q30 = defaultdict(int),defaultdict(int),defaultdict(int),defaultdict(float)
    for i in range(1,runstatsheet.nrows):
        stat = [str(j.value).strip() for j in runstatsheet.row(i)]
        sa_id = stat[0].strip().split('-')
        if len(sa_id) == 1:
            if sa_id[0] not in total_reads:
                total_reads[sa_id[0]] = int(float(stat[1]))
                human_reads[sa_id[0]] = int(float(stat[2]))
                micro_reads[sa_id[0]] = int(float(stat[4]))
                q30[sa_id[0]] = float(stat[5])
        else:
            if sa_id[0] in total_reads:
                if sa_id[1] == 'R' or sa_id[1] == 'CF':                        
                    total_reads[sa_id[0]] += int(float(stat[1]))
                    human_reads[sa_id[0]] += int(float(stat[2]))
                    micro_reads[sa_id[0]] += int(float(stat[4]))
                    q30[sa_id[0]] += float(stat[5])
            else:         
                total_reads[sa_id[0]] = int(float(stat[1]))
                human_reads[sa_id[0]] = int(float(stat[2]))
                micro_reads[sa_id[0]] = int(float(stat[4]))
                q30[sa_id[0]] = float(stat[5])

        sample[sa_id[0]].update({ 'total_reads':format(total_reads[sa_id[0]],','), 'human_reads':format(human_reads[sa_id[0]],','), 'micro_reads':format(micro_reads[sa_id[0]],','), 'q30':str('%.2f' % float((q30[sa_id[0]])/2)) })
    return sample

##模版渲染，生成报告
def getrgiTemplate(info):
    if info['tpl'].find('nj2h') > -1:
        doc = DocxTemplate(f'{str(tpl_path / "0201.nj2h.docx")}')
        doc.render(info)
        doc.save(f'{sys.argv[3]}/{info["report_id"]}_{info["department_id"]}_{info["name"]}_mNGS检测报告.docx')
    elif info['tpl'].find('boao') > -1:
        doc = DocxTemplate(f'{str(tpl_path / "0201.boao.docx")}')
        doc.render(info)
        doc.save(f'{sys.argv[3]}/{info["name"]}_{info["report_id"]}.docx')
    elif info['tpl'].find('mz') > -1:
        doc = DocxTemplate(f'{str(tpl_path / "0201.mz.docx")}')
        doc.render(info)
        doc.save(f'{sys.argv[3]}/{info["report_id"]}_{info["department_id"]}_{info["name"]}_mNGS检测报告.docx')
    elif info['tpl'].find('hy') > -1:
        doc = DocxTemplate(f'{str(tpl_path / "0201.hy.docx")}')
        doc.render(info)
        doc.save(f'{sys.argv[3]}/{info["report_id"]}_{info["department_id"]}_{info["name"]}_mNGS检测报告.docx')
    elif info['tpl'].find('fzch') > -1:
        doc = DocxTemplate(f'{str(tpl_path / "0201.fzch.docx")}')
        doc.render(info)
        doc.save(f'{sys.argv[3]}/{info["report_id"]}_{info["department_id"]}_{info["name"]}_mNGS检测报告.docx')
    else:
        if info['tpl'].find('positive2') > -1 or info['tpl'].find('negative2') > -1:
            doc = DocxTemplate(f'{str(tpl_path / "0201.zju.docx")}')
            doc.replace_pic('图片 8',f'{str(tpl_path / "zju_jcz_blank.gif")}')
            doc.replace_pic('图片 4',f'{str(tpl_path / "zju_shqz_blank.gif")}')
            doc.replace_pic('image2.png',f'{str(tpl_path / "zju_blank.png")}')
            doc.render(info)
            doc.save(f'{sys.argv[3]}/{info["report_id"]}_{info["department_id"]}_{info["name"]}_mNGS检测报告.docx')
        else:
            if info['tpl'].find('aja') > -1:
                doc = DocxTemplate(f'{str(tpl_path / "0201.aja.docx")}')
                doc.render(info)
                doc.save(f'{sys.argv[3]}/{info["report_id"]}_{info["department_id"]}_{info["name"]}_mNGS检测报告.docx')
            else:
                doc = DocxTemplate(f'{str(tpl_path / "0201.zju.docx")}')
                doc.render(info)
                doc.save(f'{sys.argv[3]}/{info["report_id"]}_{info["department_id"]}_{info["name"]}_mNGS检测报告.docx')

def main():
    sample = getSampleInfo(f'{sys.argv[1]}',f'{sys.argv[2]}')
    for k, v in sample.items():
        if k in report_sample:
            getrgiTemplate(v) 

if __name__ == '__main__':
    main()

# Tool that accumulates all .xls files in each level of a hierarchy
# into local.csv files.
# At the same time a sigma.csv is created with all at that level
# and all contents at lower levels.

# To run this, you need to
#    . bin/activate
#    pip install xlrd
#    python grind.py

import os
import fnmatch
import xlrd
import csv


def make_map(title_row):
    "return a map for column titles"
    m = {}
    for i, t in enumerate(title_row):
        m[i] = t
    return m

SF = u'SrcFile'

def remap_to_spreadsheet(data):
    samples = {}
    #get all the keys
    for row in data:
        samples.update(row)
    samples = sorted(samples)
    if SF in samples:
        samples.remove(SF)
        samples.insert(0, SF)
    titles = {}
    for i, t in enumerate(samples):
        titles[t] = i 
    # put a header row
    out = [[t for t in samples]]
    # put all the data into columns
    for row in data:
        # start with blank cells
        cells = ['' for i in range(len(titles))]

        # place whatever you encounter according to the titles map
        for k, v in row.items():
            cells[titles[k]] = unicode(v).encode('utf-8')
        out.append(cells)
    return out



def write_csv(path, file, data):
    print "  write csv %s" % str((path, file, ))
    f = os.path.join(path, file)
    data = remap_to_spreadsheet(data)
    with open(f, 'wb') as csvfile:
        writer = csv.writer(csvfile, delimiter=',',
                             quotechar="'", 
                             quoting=csv.QUOTE_MINIMAL,
                             )
        for row in data:
            writer.writerow(row)


def read_xls(path, name):
    f = os.path.join(path, name)
    wb = xlrd.open_workbook(f)
    sh = wb.sheets()[0]
    titles = sh.row(0)
    tmap = make_map([t.value for t in titles])
    outrows = []
    for rindex in range(1, sh.nrows):
        data = {}
        data[SF] = f
        for i, c in enumerate(sh.row(rindex)):
            data[tmap[i]] = c.value
        outrows.append(data)
    return outrows

ignore_dir = 'lib bin include'.split()
ignore_file = 'local sigma'.split()

def sigma(top):
    "traverse the subtree making local and sigma collections"
    print 'sigma based on %s' % top
    els = top.split('/')
    if len(els) > 1 and els[1] in ignore_dir:
        print 'skipped'
        return []
    sig = []
    for path, dirlist, filelist in os.walk(top):
        
        for di in dirlist:
            if di in ignore_dir:
                print "ignoring dir %s" % di
                continue
            sub = sigma(os.path.join(path, di))
            sig += sub

        local = []
        for name in fnmatch.filter(filelist, '*.xls'):
            if os.path.splitext(name) in ignore_file:
                print "ignoring file %s" % name
                continue
            print '  gathering %s' % name
            data = read_xls(path, name)
            local.extend(data)
            sig.extend(data)
        # for name in fnmatch.filter(filelist, '*.ods'):
        #     data = read_ods(path, name)
        #     local += data
        #     sig += data
        # for name in fnmatch.filter(filelist, '*.csv'):
        #     data = read_csv(path, name)
        #     local += data
        #     sig += data
        if sig:
            write_csv(path, 'local.csv', local)
            write_csv(path, 'sigma.csv', sig)

        break  # never recurse via walk - only via sigma call
    return sig

# set the ball rolling
sigma('.')


fix= '''
98319
98320
98321
98322
98323
98324
98325
98326
98327
98328
98329
98330
98331
98332
98333
98334
98335
98336
98337
98338











'''

fix = [x for x in fix.split("\n")]
string = ""
for row in fix:
    if row != None and row != "":
        string += row + ","

print(string[:-1])
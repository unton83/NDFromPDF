from main import NDEntity, NDList


ND_list = NDList()
ND_list.list.extend([
    NDEntity('ГОСТ 15588-86', full_name='fake name', path='fake path'),
    NDEntity('СТО 58239148-001-2006', full_name='fake name', path='fake path'),
    NDEntity('ГОСТ 379-2015', full_name='fake name', path='fake path'),
    NDEntity('ГОСТ 19903-90', full_name='fake name', path='fake path'),
    NDEntity('Серия 1.038.1-1', full_name='fake name', path='fake path'),
    NDEntity('ГОСТ 19903-90', full_name='fake name', path='another fake path'),
])

print(ND_list.labels())
# for label in ND_list:
#     full_name = get_full_name(label)
#     print(label, full_name)

ND_list.get_names()
ND_list.write_xlsx()
ND_list.write_dxf()


import pyzdde.zfileutils as zfile
a= zfile.readZRDFile('TESTRAYS.ZRD','uncompressed')
zfile.writeZRDFile(a, 'TESTRAYS_uncompressed.ZRD','uncompressed')
zfile.writeZRDFile(a, 'TESTRAYS_compressed.ZRD','compressed')
b = zfile.readZRDFile('TESTRAYS_uncompressed.ZRD','uncompressed')
c = zfile.readZRDFile('TESTRAYS_compressed.ZRD','compressed')
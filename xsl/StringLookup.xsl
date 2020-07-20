<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
  xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  version="1.0">
  <xsl:output method="text" media-type="text/plain"/>
  <!-- Row should be numeric, starts at 1, not 0; may need to add 1 to row first -->
  <xsl:param name="row" select="1015" />
  <xsl:template match="/">
    <xsl:value-of select="/sst/si[$row]/t|/s:sst/s:si[$row]/s:t"/></xsl:template>
</xsl:stylesheet>
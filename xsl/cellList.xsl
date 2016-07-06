<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
 xmlns:w="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
 <xsl:output method="text" media-type="text/plain"/>
 <xsl:template match="/">[sheet]
dimension[]="<xsl:value-of select="substring-before(//dimension/@ref|//w:dimension/@ref,':')"/>"
dimension[]="<xsl:value-of select="substring-after(//dimension/@ref|//w:dimension/@ref,':')"/>"
<xsl:apply-templates select="//c|//w:c" /></xsl:template>
 <xsl:template match="//c|//w:c">
  <xsl:if test="v|w:v">
[<xsl:value-of select="@r"/>]
row="<xsl:value-of select="../@r"/>"
val="<xsl:value-of select="v|w:v"/>"
<xsl:if test="@t">t="<xsl:value-of select="@t"/>"</xsl:if>
  </xsl:if>
 </xsl:template>
</xsl:stylesheet>
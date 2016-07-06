<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
  xmlns:o="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  version="1.0">
  <xsl:output method="text" media-type="text/plain"/>
  <xsl:param name="page" select="'xl/worksheets/sheet'" />
  <xsl:param name="r" select="'true'"/>
  <xsl:template match="/"><xsl:apply-templates select="o:workbook|workbook" />
 </xsl:template>
  <xsl:template  match="o:workbook|workbook">[worksheets]<xsl:apply-templates select="o:sheets|sheets" /></xsl:template>
  <xsl:template match="o:sheets|sheets">
    <xsl:for-each select="o:sheet|sheet">
title[]="<xsl:value-of select="@name"/>"
<xsl:choose>
 <xsl:when test="$r='true'">
xml[]="<xsl:value-of select="$page"/><xsl:value-of select="substring-after(@r:id, 'rId')"/>.xml"   
  </xsl:when>
  <xsl:otherwise>
xml[]="<xsl:value-of select="$page"/><xsl:value-of select="@sheetId"/>.xml"
  </xsl:otherwise>
</xsl:choose>  
  </xsl:for-each>
 </xsl:template>
</xsl:stylesheet>
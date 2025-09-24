<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
  xmlns:m="http://www.w3.org/1998/Math/MathML"
  xmlns:mml="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <!-- Entrada principal -->
  <xsl:template match="m:math">
    <mml:oMathPara>
      <mml:oMath>
        <xsl:apply-templates select="*"/>
      </mml:oMath>
    </mml:oMathPara>
  </xsl:template>

  <!-- Identificadores, números y operadores -->
  <xsl:template match="m:mi|m:mo|m:mn|m:mtext">
    <mml:r>
      <w:rPr><w:rFonts ascii="Cambria Math"/></w:rPr>
      <mml:t><xsl:value-of select="."/></mml:t>
    </mml:r>
  </xsl:template>

  <!-- Filas (agrupaciones) -->
  <xsl:template match="m:mrow">
    <xsl:apply-templates select="*"/>
  </xsl:template>

  <!-- Fracciones -->
  <xsl:template match="m:mfrac">
    <mml:f>
      <mml:num><xsl:apply-templates select="*[1]"/></mml:num>
      <mml:den><xsl:apply-templates select="*[2]"/></mml:den>
    </mml:f>
  </xsl:template>

  <!-- Superíndices -->
  <xsl:template match="m:msup">
    <mml:sSup>
      <mml:e><xsl:apply-templates select="*[1]"/></mml:e>
      <mml:sup><xsl:apply-templates select="*[2]"/></mml:sup>
    </mml:sSup>
  </xsl:template>

  <!-- Subíndices -->
  <xsl:template match="m:msub">
    <mml:sSub>
      <mml:e><xsl:apply-templates select="*[1]"/></mml:e>
      <mml:sub><xsl:apply-templates select="*[2]"/></mml:sub>
    </mml:sSub>
  </xsl:template>

  <!-- Sub + super -->
  <xsl:template match="m:msubsup">
    <mml:sSubSup>
      <mml:e><xsl:apply-templates select="*[1]"/></mml:e>
      <mml:sub><xsl:apply-templates select="*[2]"/></mml:sub>
      <mml:sup><xsl:apply-templates select="*[3]"/></mml:sup>
    </mml:sSubSup>
  </xsl:template>

  <!-- Raíz cuadrada -->
  <xsl:template match="m:msqrt">
    <mml:rad>
      <mml:deg/>
      <mml:e><xsl:apply-templates select="*"/></mml:e>
    </mml:rad>
  </xsl:template>

  <!-- Raíz enésima -->
  <xsl:template match="m:mroot">
    <mml:rad>
      <mml:deg><xsl:apply-templates select="*[2]"/></mml:deg>
      <mml:e><xsl:apply-templates select="*[1]"/></mml:e>
    </mml:rad>
  </xsl:template>

  <!-- Ignorar texto puro fuera de nodos -->
  <xsl:template match="text()"/>

</xsl:stylesheet>

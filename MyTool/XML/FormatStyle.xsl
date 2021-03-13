<?xml version="1.0" ?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">

  <xsl:output method="xml" />

  <xsl:template match="/">
    <xsl:apply-templates select="NewDataSet" />
  </xsl:template>

  <xsl:template match="/NewDataSet" >
        
    <xsl:element name="xml">
      <xsl:attribute name="version">1.0</xsl:attribute>
      <xsl:attribute name="encoding">UTF-8</xsl:attribute>?
    </xsl:element>

    <doi_batch version="4.1.0" xmlns="http://www.crossref.org/schema/4.1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.crossref.org/schema/4.1.0 crossref4.1.0.xsd">

      <xsl:element name="head">
        <xsl:element name="doi_batch_id">doi_12345</xsl:element>
        <xsl:element name="timestamp">20011003110604</xsl:element>
        <xsl:element name="depositor">
          <xsl:element name="name">Jaypee Brothers Medical Publishers (P) ltd</xsl:element>
          <xsl:element name="email_address">jaypee@jaypeebrothers.com</xsl:element>
        </xsl:element>
        <xsl:element name="registrant">Jaypee Brothers Medical Publishers (P) ltd</xsl:element>
      </xsl:element>



      <xsl:element name="body">
        <xsl:for-each select="Table">

          <xsl:element name="book">
            <xsl:attribute name="book_type">
              <xsl:value-of select="eBOOKTYPE"/>
            </xsl:attribute>

            <xsl:element name="book_metadata">
              <xsl:attribute name="language">
                <xsl:value-of select="LANGUAGE"/>
              </xsl:attribute>


              <xsl:element name="contributors">
                <xsl:element name="person_name">
                  <xsl:attribute name="sequence">First</xsl:attribute>
                  <xsl:attribute name="contributor_role">Editor</xsl:attribute>
                  <xsl:element name="given_name">
                    <xsl:value-of select="AUTHORNAME"/>
                  </xsl:element>
                  <xsl:element name="surname">
                    <xsl:value-of select="SURNAME"/>
                  </xsl:element>
                </xsl:element>

                <xsl:element name="person_name">
                  <xsl:attribute name="sequence">Additional</xsl:attribute>
                  <xsl:attribute name="contributor_role">Editor</xsl:attribute>
                  <xsl:element name="given_name">
                    <xsl:value-of select="AUTHORNAME"/>
                  </xsl:element>
                  <xsl:element name="surname">
                    <xsl:value-of select="SURNAME"/>
                  </xsl:element>
                </xsl:element>
              </xsl:element>

              <xsl:element name="titles">
                <xsl:element name="title">
                  <xsl:value-of select="TITLE"/>
                </xsl:element>
              </xsl:element>

              <xsl:element name="volume">
                <xsl:value-of select="VOLUME"/>
              </xsl:element>

              <xsl:element name="edition_number">
                <xsl:value-of select="EDITION"/>
              </xsl:element>

              <xsl:element name="publication_date">
                <xsl:attribute name="media_type">print</xsl:attribute>
                <xsl:element name="month">10</xsl:element>
                <xsl:element name="day">5</xsl:element>
                <xsl:element name="year">2009</xsl:element>
              </xsl:element>

              <xsl:element name="isbn">
                <xsl:value-of select="ISBN"/>
              </xsl:element>

              <xsl:element name="publisher">
                <xsl:element name="publisher_name">
                  <xsl:value-of select="PUBLISHER"/>
                </xsl:element>
                <xsl:element name="publisher_place">New Delhi</xsl:element>
              </xsl:element>

              <xsl:element name="doi_data">
                <xsl:element name="doi">10.9999/0-19-262706-6</xsl:element>
                <xsl:element name="resource">http://www.oup.co.uk/isbn/0-19-262706-6</xsl:element>
              </xsl:element>
            </xsl:element>

            <xsl:element name="content_item">
              <xsl:attribute name="component_type">
                <xsl:value-of select="TYPE"/>
              </xsl:attribute>
              <xsl:attribute name="level_sequence_number">
                <xsl:value-of select="SNO"/>
              </xsl:attribute>
              <xsl:attribute name="publication_type">
                <xsl:value-of select="DISPLAYTEXT"/>
              </xsl:attribute>
              <xsl:element name="titles">
                <xsl:element name="title">
                  <xsl:value-of select="TITLE"/>
                </xsl:element>
                <xsl:element name="doi_data">
                  <xsl:element name="doi">
                    10.9999/0-19-262706-6.<xsl:value-of select="TITLE"/>
                  </xsl:element>
                  <xsl:element name="resource">
                    http://www.oup.co.uk/isbn/0-19-262706-6/<xsl:value-of select="TITLE"/>
                  </xsl:element>
                </xsl:element>
              </xsl:element>
            </xsl:element>

          </xsl:element>

        </xsl:for-each>
      </xsl:element>

    </doi_batch>
    
  </xsl:template>
</xsl:transform>
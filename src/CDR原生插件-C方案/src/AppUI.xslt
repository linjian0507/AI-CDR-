<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:frmwrk="Corel Framework Data">
  <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

  <frmwrk:uiconfig>
    <frmwrk:applicationInfo userConfiguration="true" />
  </frmwrk:uiconfig>

  <!-- 原样复制所有节点 -->
  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- 注册 UI 项: 工具栏按钮 + 停靠窗承载的 WPF 控件 -->
  <xsl:template match="uiConfig/items">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>

      <!-- 工具栏按钮: 切换 AI矢量助手 v2.3.7 停靠窗 -->
      <itemData guid="1f6f1448-5111-4dc2-b84f-f386b5f194ac" noBmpOnMenu="true"
          type="checkButton"
          check="*Docker('97d37334-68f4-40aa-be6f-7677f35de4c7')"
          userCaption="AI矢量助手 v2.3.7"
          enable="true">
      </itemData>

      <!-- 停靠窗里承载的 WPF 控件(进程内加载) -->
      <itemData guid="3566aeac-c7ef-4927-a00e-2b067f809fe0"
          type="wpfhost"
          hostedType="Addons\AIVectorHelper\AIVectorHelper.dll,AIVectorHelper.MainPanel"
          enable="true">
      </itemData>
    </xsl:copy>
  </xsl:template>

  <!-- 定义工具栏, 放入上面的按钮 -->
  <xsl:template match="uiConfig/commandBars">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <commandBarData guid="9e10c5b3-bd16-4d2c-9bde-920ec67a5b19"
                      nonLocalizableName="AIVectorHelper"
                      userCaption="AI矢量助手 v2.3.7"
                      locked="false"
                      type="toolbar">
        <toolbar>
          <item guidRef="1f6f1448-5111-4dc2-b84f-f386b5f194ac" dock="top"/>
        </toolbar>
      </commandBarData>
    </xsl:copy>
  </xsl:template>

  <!-- 把工具栏挂到主界面顶部停靠区(这些 GUID 是 CDR 框架内置的, 各插件通用) -->
  <xsl:template match="uiConfig/containers/container[@guid='bee85f91-3ad9-dc8d-48b5-d2a87c8b2109']/container[@guid='Framework_MainFrame-layout']/dockHost[@guid='894bf987-2ec1-8f83-41d8-68f6797d0db4']/toolbar[@guidRef='c2b44f69-6dec-444e-a37e-5dbf7ff43dae']">
    <xsl:copy-of select="."/>
    <toolbar guidRef="9e10c5b3-bd16-4d2c-9bde-920ec67a5b19" dock="top" />
  </xsl:template>

  <!-- 定义停靠窗, 把 WPF 控件填进去 -->
  <xsl:template match="uiConfig/dockers">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <dockerData guid="97d37334-68f4-40aa-be6f-7677f35de4c7"
                  userCaption="AI矢量助手 v2.3.7"
                  wantReturn="true"
                  focusStyle="noThrow">
        <container>
          <item dock="fill" margin="0,0,0,0" guidRef="3566aeac-c7ef-4927-a00e-2b067f809fe0"/>
        </container>
      </dockerData>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>

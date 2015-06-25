<xsl:stylesheet version = '1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>
<xsl:output method = "xml" encoding = "UTF-8"/>
<xsl:variable name="CDATA1"><![CDATA[<![CDATA[
]]></xsl:variable>
<xsl:variable name="CDATA21"><![CDATA[
]]]]></xsl:variable>
<xsl:variable name="CDATA22">></xsl:variable>

<xsl:template match='/'>
<xsl:param name="lev"><xsl:value-of select="/root/@indent"/></xsl:param>
	<testsuite>
		<xsl:attribute name="name">
			<xsl:value-of select="/root/@title"/>
		</xsl:attribute>
		<xsl:attribute name="t1">
			<xsl:value-of select="count(/root/testnode[./@indent = /root/testnode[1]/@indent])"/>
		</xsl:attribute>
		<xsl:for-each select="/root/testnode[./@indent = /root/testnode[1]/@indent]">
			<xsl:call-template name="testsuite_or_testcase">
				<!-- такой уровень у текущей корневой записи, вложенные записи будут искать все записи такого уровня,
					а среди них выбирать ту, в которую они "вложены".
					Текущий элемент является конневым. Понятно, что все элементы вложены в него, поэтому для /root будет спецобработка.	-->
				<xsl:with-param name="prevlev">
					<xsl:value-of select="$lev"/>
				</xsl:with-param>
				<!-- Это число задаёт тот уровень, элементы имеющие который будут искаться как элементы, вложенные в текущий элемент.
					Шаблон testsuite_or_testcase толком не знает, что ему надо найти. Парамер $lev сообщит ему, что нужны все элементы с уровнем равным $lev. -->
				<xsl:with-param name="lev">
					<xsl:value-of select="/root/testnode[1]/@indent"/>
				</xsl:with-param>
				<!-- Этот параметр позволит вложенным записям сравнить ближайшую к ним запись, находящуюся выше их, и имеющую уровень их параметра prevlev с настоящей родительской записью - текущим параметром.
					Если эта ближайшая сверху запись окажется не родительской, то обработка записи не будет осуществляться. -->
				<xsl:with-param name="parentNode">
					<xsl:value-of select="generate-id(current())"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:for-each>
	</testsuite>
</xsl:template>

<!-- Определение того, чем является текущий элемент testnode - группой тестов или тестом -->
<xsl:template name="testsuite_or_testcase">
	<xsl:param name="prevlev"></xsl:param>
	<xsl:param name="lev"></xsl:param>
	<xsl:param name="parentNode"></xsl:param>
	<!--
	<debug>
		<xsl:attribute name="title">
			<xsl:value-of select="@title"/>
		</xsl:attribute>
	</debug>
	-->
	<!-- если ближайшая верхняя запись верхнего уровня является родительской, или если родительская запись - корневой элемент,
		то текущая запись будет обработана, иначе обработка текущей записи будет пропущена -->
	<xsl:if test="(generate-id(preceding-sibling::testnode[@indent = $prevlev][position()=1]) = $parentNode) or (count(preceding-sibling::testnode[@indent = $prevlev][position()=1]) = 0)">
		<!--если	у следующего элемента testnode уровень <= уровню текущего элемента testnode
			или	следующего элемента testnode просто нет (текущий элемент последний)
			то	текущий элемент - тест (не группа тестов)-->
		<xsl:if test="(not(following-sibling::testnode[1]/@indent > ./@indent)) or (not(following-sibling::testnode[1]))">
			<xsl:call-template name="testcase">
			</xsl:call-template>
		</xsl:if>
		
		<!--если	у следующего элемента testnode уровень > уровня текущего элемента testnode,
				и	этот следующий элемент существует,
				то следeющий элемент вложен в текущий, а значит текущий - группа тестов	-->
		<xsl:if test="(following-sibling::testnode[1]/@indent > ./@indent) and (following-sibling::testnode[1])">
			<xsl:call-template name="testsuite">
				<xsl:with-param name="prevlev">
					<xsl:value-of select="./@indent"/>
				</xsl:with-param>
				<xsl:with-param name="lev">
					<xsl:value-of select="following-sibling::testnode[1]/@indent"/>
				</xsl:with-param>
				<xsl:with-param name="parentNode">
					<xsl:value-of select="generate-id(current())"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:if>
</xsl:template>

<!-- Обработка текущего элемента testnode как группы тестов -->
<xsl:template name="testsuite">
	<xsl:param name="prevlev"></xsl:param>
	<xsl:param name="lev"></xsl:param>
	<xsl:param name="parentNode"></xsl:param>
	<testsuite>
		<xsl:attribute name="name">
			<xsl:call-template name="testName"/>
		</xsl:attribute>
		<xsl:call-template name="node_order"/>
		<xsl:call-template name="details"/>
		
		<xsl:for-each select="following-sibling::testnode[./@indent = $lev]">
			<xsl:call-template name="testsuite_or_testcase">
				<xsl:with-param name="prevlev">
					<xsl:value-of select="$prevlev"/>
				</xsl:with-param>
				<xsl:with-param name="lev">
					<xsl:value-of select="$lev"/>
				</xsl:with-param>
				<xsl:with-param name="parentNode">
					<xsl:value-of select="$parentNode"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:for-each>
	</testsuite>
</xsl:template>

<!-- Обработка текущего элемента testnode как теста -->
<xsl:template name="testcase">
	<testcase>
		<xsl:attribute name="name">
			<xsl:call-template name="testName"/>
		</xsl:attribute>

		<xsl:call-template name="node_order"/>
		<externalid><xsl:value-of select="generate-id(current())"/></externalid>
		<version><![CDATA[1]]></version>
		<xsl:call-template name="summary"/>
		<preconditions><![CDATA[]]></preconditions>
		<!--<execution_type><![CDATA[1]]></execution_type>-->
		<!--<importance><![CDATA[2]]></importance>-->

		<xsl:if test="count(testSteps) > 0">
			<steps>
				<xsl:for-each select="testSteps/step">
					<step>
					<step_number>
						<!--<xsl:value-of select = "$CDATA1" disable-output-escaping = "yes" />-->
						<xsl:value-of select = "position()" />
						<!--<xsl:value-of select = "concat($CDATA21,$CDATA22)" disable-output-escaping = "yes" />-->
					</step_number>
					<actions>
						<xsl:value-of select = "$CDATA1" disable-output-escaping = "yes" />
						<xsl:copy-of select="text/*" />
						<xsl:value-of select = "concat($CDATA21,$CDATA22)" disable-output-escaping = "yes" />
					</actions>

					<expectedresults>
						<xsl:value-of select = "$CDATA1" disable-output-escaping = "yes" />
						<xsl:copy-of select="result/*" />
						<xsl:value-of select = "concat($CDATA21,$CDATA22)" disable-output-escaping = "yes" />
					</expectedresults>
											
					<!--<execution_type><![CDATA[1]]></execution_type>-->
					</step>
				</xsl:for-each>
			</steps>
		</xsl:if>
	</testcase>
</xsl:template>


<xsl:template name='template1'>
<xsl:param name="prevlev"></xsl:param>
<xsl:param name="lev"></xsl:param>
<xsl:param name="parentNode"></xsl:param>
	<xsl:for-each select="following-sibling::testnode[@indent = $lev]">
		<xsl:if test="(preceding-sibling::testnode[@indent = $prevlev][position()=1]) = $parentNode">
			<xsl:if test="following-sibling::testnode[1]/@indent > $lev">
				<testsuite>
					<xsl:attribute name="name">
						<xsl:call-template name="testName"/>
					</xsl:attribute>
					<xsl:call-template name="node_order"/>
					<xsl:call-template name="details"/>
					<xsl:call-template name="template1">
						<xsl:with-param name="prevlev">
							<xsl:value-of select="$lev"/>
						</xsl:with-param>
						<xsl:with-param name="lev">
							<xsl:value-of select="following-sibling::testnode[1]/@indent"/>
						</xsl:with-param>
						<xsl:with-param name="parentNode">
							<xsl:value-of select="current()"/>
						</xsl:with-param>
					</xsl:call-template>
				</testsuite>
			</xsl:if>
			<xsl:if test="(not(following-sibling::testnode[1]/@indent > $lev)) or (not(following-sibling::testnode[1]) or (following-sibling::testnode[1] and not(following-sibling::testnode[1][./@indent])))">
				<testcase>
					<xsl:attribute name="name">
						<xsl:call-template name="testName"/>
					</xsl:attribute>
					<xsl:call-template name="node_order"/>
					<externalid><![CDATA[0]]></externalid>
					<version><![CDATA[1]]></version>
					<xsl:call-template name="summary"/>
					<preconditions><![CDATA[]]></preconditions>
					<execution_type><![CDATA[1]]></execution_type>
					<importance><![CDATA[2]]></importance>
					<xsl:if test="count(testSteps) > 0">
						<steps>
							<xsl:for-each select="testSteps/step">
								<step>
									<step_number>
										<!--<xsl:value-of select = "$CDATA1" disable-output-escaping = "yes" />-->
										<xsl:value-of select = "position()" />
										<!--<xsl:value-of select = "concat($CDATA21,$CDATA22)" disable-output-escaping = "yes" />-->
									</step_number>
									<actions>
										<xsl:value-of select = "$CDATA1" disable-output-escaping = "yes" />
										<xsl:copy-of select="text/*" />
										<xsl:value-of select = "concat($CDATA21,$CDATA22)" disable-output-escaping = "yes" />
									</actions>

									<expectedresults>
										<xsl:value-of select = "$CDATA1" disable-output-escaping = "yes" />
										<xsl:copy-of select="result/*" />
										<xsl:value-of select = "concat($CDATA21,$CDATA22)" disable-output-escaping = "yes" />
									</expectedresults>
									
									<execution_type><![CDATA[1]]></execution_type>
								</step>
							</xsl:for-each>
						</steps>
					</xsl:if>
				</testcase>
			</xsl:if>
		</xsl:if>
	</xsl:for-each>
</xsl:template>

<xsl:template name='info'>
	<xsl:value-of select = "$CDATA1" disable-output-escaping = "yes" />
	<p style="text-align:center">
		<strong>
			<xsl:value-of select="normalize-space(@title)"/>
		</strong>
	</p>
	<xsl:copy-of select="def/*" />
	<xsl:value-of select = "concat($CDATA21,$CDATA22)" disable-output-escaping = "yes" />
</xsl:template>


<xsl:template name='details'>
	<details>
		<xsl:call-template name="info"/>
	</details>
</xsl:template>

<xsl:template name='node_order'>
	<node_order>
		<xsl:value-of select="position()"/>
	</node_order>
</xsl:template>

<xsl:template name='summary'>
	<summary>
		<xsl:call-template name="info"/>
	</summary>
</xsl:template>

<xsl:template name='testName'>
	<xsl:value-of select="substring(@title, 1, 100)"/>
</xsl:template>

</xsl:stylesheet>
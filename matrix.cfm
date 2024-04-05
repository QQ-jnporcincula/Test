<!--- for key facts --->
<cfoutput>
<cfset tm_for_kf = 0>
</cfoutput>

<CFQUERY name="GetQprime" dataSource="fbapp">
SELECT Log([k])/Log(10)+2*Log([loo])/Log(10) AS Q, POPGROWTH.*
FROM POPGROWTH
WHERE 	(POPGROWTH.SpecCode=#SpecID#) 													and
		(POPGROWTH.LinfLmax=0)														and
		(popgrowth.type is not null)												and
		(popgrowth.auxim <> 'Doubtful')
ORDER BY Log([k])/Log(10)+2*Log([loo])/Log(10);
</cfquery>

<cfif #getqprime.recordcount# is 0>
	<cfset Go4 = 0>
	<cfif #withmaturity.recordcount# is not 0>
		<CFQUERY name="TM01" datasource="fbapp">
		SELECT Avg(MATURITY.tm) AS AvgOftm
		FROM MATURITY
		WHERE (	((MATURITY.Sex)<>'male') 							AND
			((MATURITY.tm)<>0 And (MATURITY.tm) Is Not Null)	AND
			((MATURITY.Speccode)=#SpecID#));	
		</CFQUERY>
		<cfif #tm01.avgoftm# GT 0>
			<cfoutput>
			<cfset tm_for_kf = #tm01.avgoftm#>
			</cfoutput>				
		<cfelse>
			<CFQUERY name="TM02" datasource="fbapp">	
			SELECT Avg(([agematmin]+[agematmin2])/2) AS vAvg
			FROM MATURITY
			WHERE 	(((MATURITY.AgeMatMin)<>0 And (MATURITY.AgeMatMin) Is Not Null) AND
			((MATURITY.AgeMatMin2)<>0 And (MATURITY.AgeMatMin2) Is Not Null) AND
			((MATURITY.Sex)<>'male') AND ((MATURITY.Speccode)=#SpecID#));
			</cfquery>			
			<cfif #tm02.vAvg# GT 0>
				<cfoutput>
				<cfset tm_for_kf = #tm02.vAvg#>
				</cfoutput>		
			<cfelse>
				<CFQUERY name="TM03" datasource="fbapp">
				SELECT Avg(MATURITY.agematmin) AS AvgOfagematmin
				FROM MATURITY
				WHERE (	((MATURITY.Sex)<>'male') 							AND
				((MATURITY.agematmin)<>0 And (MATURITY.agematmin) Is Not Null)	AND
				((MATURITY.Speccode)=#SpecID#));	
				</CFQUERY>		
				<cfif #tm03.Avgofagematmin# GT 0>
					<cfoutput>
					<cfset tm_for_kf = #tm03.Avgofagematmin#>
					</cfoutput>					
				<cfelse>
					<CFQUERY name="TM04" datasource="fbapp">
					SELECT Avg(POPCHAR.tmax) AS AvgOftmax						
					FROM POPCHAR
					WHERE (	((POPCHAR.Sex)<>'male') 							AND
					((POPCHAR.tmax)<>0 And (POPCHAR.tmax) Is Not Null)	AND
					((POPCHAR.Speccode)=#SpecID#));								
					</CFQUERY>		
					<cfif #tm04.Avgoftmax# GT 0>
						<cfoutput>
						<cfset tm_for_kf = #tm04.Avgoftmax#>
						<cfset Go4 = 1>
						</cfoutput>					
					<cfelse>
						<cfoutput>
						<cfset tm_for_kf = 0>
						</cfoutput>									
					</cfif>			
				</cfif>			
			</cfif>
		</cfif>
	<cfelse>	
		<CFQUERY name="TM04" datasource="fbapp">
		SELECT Avg(POPCHAR.tmax) AS AvgOftmax						
		FROM POPCHAR
		WHERE (	((POPCHAR.Sex)<>'male') 							AND
				((POPCHAR.tmax)<>0 And (POPCHAR.tmax) Is Not Null)	AND
				((POPCHAR.Speccode)=#SpecID#));								
		</CFQUERY>		
		<cfif #tm04.Avgoftmax# GT 0>
			<cfoutput>
			<cfset tm_for_kf = #tm04.Avgoftmax#>
			<cfset Go4 = 1>		
			</cfoutput>					
		<cfelse>
			<cfoutput>
			<cfset tm_for_kf = 0>
			</cfoutput>									
		</cfif>			
	</cfif>	
</cfif>

<!--- for key facts --->


<cfif #getqprime.recordcount# is not 0>		
	<a href="/PopDyn/KeyfactsSummary_1.cfm?ID=#URLEncodedFormat(SpecID)#
	&genusname=#urlencodedformat(GenusN)#
	&speciesname=#urlencodedformat(SpeciesN)#
	&vstockcode=#urlencodedformat(getstockcode.stockcode)#		
	&fc=#detail.detailfield22#">		
	<!--- Key facts* --->
<cfelse>
	<cfif #tm_for_kf# is 0>			
		<a href="/PopDyn/KeyfactsSummary_2v2.cfm?ID=#URLEncodedFormat(SpecID)#
		&genusname=#urlencodedformat(GenusN)#
		&speciesname=#urlencodedformat(SpeciesN)#
		&vstockcode=#urlencodedformat(getstockcode.stockcode)#			
		&fc=#detail.detailfield22#">		
		<!--- Key facts****--->
	<cfelse>
		<cfif #Go4# is 0>
			<a href="/PopDyn/KeyfactsSummary_3.cfm?ID=#URLEncodedFormat(SpecID)#
			&genusname=#urlencodedformat(GenusN)#
			&speciesname=#urlencodedformat(SpeciesN)#
			&vstockcode=#urlencodedformat(getstockcode.stockcode)#			
			&fc=#detail.detailfield22#
			&var_tm=#tm_for_kf#">				
			<!--- Key facts**--->
		<cfelse>						
			<a href="/PopDyn/KeyfactsSummary_4.cfm?ID=#URLEncodedFormat(SpecID)#
			&genusname=#urlencodedformat(GenusN)#
			&speciesname=#urlencodedformat(SpeciesN)#
			&vstockcode=#urlencodedformat(getstockcode.stockcode)#			
			&fc=#detail.detailfield22#
			&var_tmax=#tm_for_kf#">				
			<!--- Key facts***--->
		</cfif>			
	</cfif>
</cfif>		


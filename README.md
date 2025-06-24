### Execution Instructions

## Run All Scenarios
mvn clean test

## Combine Multiple Tags (AND logic)
mvn clean test -Dcucumber.filter.tags="@happy-path and @edge-case and @filter and @pagination"

## Prerequisites
Java 21+ (JDK installed and JAVA_HOME set)  
Maven 3.8+ (installed and on your PATH)  
Internet access (for downloading WebDriverManager binaries)

## Reporting
Extent Report: target/NextGen_BDDResults/LatestResults/Reports_<date>/DetailedReport/Detailed Report.html
Screenshots: target/NextGen_BDDResults/LatestResults/Reports_<date>/DetailedReport/Screenshots
Logs: target/NextGen_BDDResults/LatestResults/Reports_<date>/Logs/NhsbsaRunner_Chrome_Test.log

## Feature File Path
NextGenBDD/src/test/java/com/nhsbsa/features/nhs_job_search.feature

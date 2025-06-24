Feature: NHS Jobs Search
  In order to find relevant NHS job vacancies
  As a job-seeker
  I want to search, filter, sort and page through results on the NHS Jobs site

  Background: 
    Given I am on the NHS Jobs search page
      | url | https://www.jobs.nhs.uk/candidate/search |

  @happy-path
  Scenario Outline: Search returns only matching jobs sorted by newest first
    When I search for jobs with:
      | keyword   | location   | distance   | employer   | payrange   |
      | <keyword> | <location> | <distance> | <employer> | <payrange> |
    Then I should see only jobs matching "<keyword>"
    When I sort my search results with the "Date Posted (newest)"
    Then results should be sorted by newest date posted

    Examples: 
      | keyword      | location | distance  | employer                        | payrange           |
      | Physiologist | London   | +10 Miles | NHS Business Services Authority | £40,000 to £50,000 |
      | Audiologist  | London   | +10 Miles | NHS Business Services Authority | £40,000 to £50,000 |

  @edge-case
  Scenario: No results when filters are too restrictive
    When I search for jobs with:
      | keyword   | location | distance  | employer                        | payrange           |
      | Astronaut | London   | +10 Miles | NHS Business Services Authority | £40,000 to £50,000 |
    Then I should see a No result found message

  @filter
  Scenario Outline: Filter by employer only
    When I search for jobs with:
      | keyword   | location   | distance   | employer   | payrange   |
      | <keyword> | <location> | <distance> | <employer> | <payrange> |
    Then I should see only jobs from employer "<employer>"

    Examples: 
      | keyword | location  | distance | employer                        | payrange |
      |         | Fleetwood | +5 Miles | NHS Business Services Authority |          |

  @pagination
  Scenario: Pager navigates through all search-result pages
    When I search for jobs with:
      | keyword | location | distance  | employer                        | payrange           |
      | Nurse   | London   | +50 Miles | NHS Business Services Authority | £30,000 to £40,000 |
    Then I should be able to traverse to each Next page until there are none

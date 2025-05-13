
-- ==============================================
-- SQL Script to Mirror Excel VBA Functionality
-- For: MGS617 Final Project - Gambling Data
-- ==============================================

-- 1. StateData: Monthly gambling data per state
CREATE TABLE StateData (
    State VARCHAR(50),
    DateKey DATE,
    Handle DECIMAL(18, 2),
    Revenue DECIMAL(18, 2),
    WinLoss DECIMAL(18, 2),
    Taxes DECIMAL(18, 2)
);

-- 2. PopulationCensus: Annual population estimates
CREATE TABLE PopulationCensus (
    State VARCHAR(50),
    Year INT,
    Population BIGINT
);

-- 3. StateCoordinates: Latitude and longitude of state centers
CREATE TABLE StateCoordinates (
    State VARCHAR(50),
    Latitude FLOAT,
    Longitude FLOAT
);

-- ==============================================
-- Queries
-- ==============================================

-- A. Total Revenue and Handle by State and Year (Line Chart)
SELECT 
    State,
    EXTRACT(YEAR FROM DateKey) AS Year,
    SUM(Handle) AS TotalHandle,
    SUM(Revenue) AS TotalRevenue
FROM StateData
GROUP BY State, EXTRACT(YEAR FROM DateKey)
ORDER BY State, Year;

-- B. Revenue vs. Population (Correlation Line Chart)
SELECT 
    sd.State,
    EXTRACT(YEAR FROM sd.DateKey) AS Year,
    SUM(sd.Revenue) AS TotalRevenue,
    pc.Population
FROM StateData sd
JOIN PopulationCensus pc
  ON sd.State = pc.State AND EXTRACT(YEAR FROM sd.DateKey) = pc.Year
GROUP BY sd.State, EXTRACT(YEAR FROM sd.DateKey), pc.Population
ORDER BY sd.State, Year;

-- C. Handle Per Capita by State (Bubble Map Data)
WITH AvgHandle AS (
    SELECT 
        State,
        AVG(Handle) AS AvgMonthlyHandle
    FROM StateData
    GROUP BY State
),
AvgPop AS (
    SELECT 
        State,
        AVG(Population) AS AvgPopulation
    FROM PopulationCensus
    GROUP BY State
)
SELECT 
    ah.State,
    (ah.AvgMonthlyHandle / ap.AvgPopulation) AS HandlePerCapita,
    sc.Latitude,
    sc.Longitude
FROM AvgHandle ah
JOIN AvgPop ap ON ah.State = ap.State
JOIN StateCoordinates sc ON ah.State = sc.State
ORDER BY HandlePerCapita DESC;

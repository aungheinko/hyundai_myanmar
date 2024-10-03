# Let's calculate the daily, weekly, and yearly earnings based on a 10% APR for 1000 USDT.

principal = 1000  # Initial investment in USDT
annual_rate = 0.10  # 10% APR
days_in_year = 365  # Assuming a non-leap year

# Daily earnings
daily_rate = annual_rate / days_in_year
daily_earnings = principal * daily_rate

# Weekly earnings
weekly_earnings = daily_earnings * 7

# Yearly earnings
yearly_earnings = principal * annual_rate

daily_earnings, weekly_earnings, yearly_earnings

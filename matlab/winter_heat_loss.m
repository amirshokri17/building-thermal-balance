% MATLAB script to calculate heat loss using overall building dimensions

% Step 1: Input shared building dimensions and temperatures
fprintf('--- Building Dimensions and Conditions ---\n');
b_length = input('Enter building length (m): ');
b_width = input('Enter building width (m): ');
b_height = input('Enter building height (m): ');

U_wall = input('Enter U-value for walls (W/m^2K): ');
U_floor = input('Enter U-value for floor (W/m^2K): ');
U_ceil = input('Enter U-value for ceiling (W/m^2K): ');

fprintf('\n--- f_x values for each wall ---\n');
fx_wall_SW = input('Enter f_x for SW wall (width side): ');
fx_wall_SE = input('Enter f_x for SE wall (length side): ');
fx_wall_NE = input('Enter f_x for NE wall (width side): ');
fx_wall_NW = input('Enter f_x for NW wall (length side): ');

fx_floor = input('Enter f_x for floor: ');
fx_ceil = input('Enter f_x for ceiling: ');

Tin = input('Enter internal temperature (°C): ');
Tout = input('Enter external temperature for walls and ceiling (°C): ');
T_ground = input('Enter ground temperature for floor (°C): ');

% Calculate ∆T
deltaT_wall = Tin - Tout;
deltaT_floor = Tin - T_ground;

% Step 2: Calculate wall areas
area_wall_side = b_width * b_height;
area_wall_front = b_length * b_height;

% Step 3: Window calculation
fprintf('\n--- Window Heat Loss Calculation ---\n');
n_windows = input('Enter number of windows: ');
win_width = input('Enter window horizontal dimension (m): ');
win_height = input('Enter window vertical dimension (m): ');
U_window = input('Enter U-value for windows (W/m^2K): ');
fx_window = input('Enter f_x for windows: ');

area_window_total = n_windows * win_width * win_height;
Q_window = deltaT_wall * area_window_total * U_window * fx_window;

% Step 4: Subtract window area from SW wall and recalculate its heat loss
area_wall_SW_net = area_wall_side - area_window_total;
Q_wall_SW = deltaT_wall * area_wall_SW_net * U_wall * fx_wall_SW;

% Step 5: Remaining walls
Q_wall_SE = deltaT_wall * area_wall_front * U_wall * fx_wall_SE;
Q_wall_NE = deltaT_wall * area_wall_side * U_wall * fx_wall_NE;
Q_wall_NW = deltaT_wall * area_wall_front * U_wall * fx_wall_NW;

% Step 6: Floor and ceiling
area_floor = b_width * b_length;
area_ceil = area_floor;
Q_floor = deltaT_floor * area_floor * U_floor * fx_floor;
Q_ceil = deltaT_wall * area_ceil * U_ceil * fx_ceil;

% Step 7: Air exchange heat loss
fprintf('\n--- Air Exchange Heat Loss Calculation ---\n');
n_air = input('Enter air exchange rate (ACH): ');
f = 1/3600;
rho = 1.25;
c_p = 1000;
volume = b_length * b_width * b_height;
Q_tot_air_exchange = n_air * volume * f * rho * c_p * deltaT_wall;

% Step 8: Total heat loss
Q_tot_heat_transfer = Q_wall_SW + Q_wall_SE + Q_wall_NE + Q_wall_NW + Q_floor + Q_ceil + Q_window;
Q_tot_final = Q_tot_heat_transfer + Q_tot_air_exchange;

% Step 9: Display results
fprintf('\n%-15s %-12s %-10s\n','Element','Area(m^2)','Q_loss(W)');
fprintf('%-15s %-12.2f %-10.2f\n','Wall SW', area_wall_SW_net, Q_wall_SW);
fprintf('%-15s %-12.2f %-10.2f\n','Wall SE', area_wall_front, Q_wall_SE);
fprintf('%-15s %-12.2f %-10.2f\n','Wall NE', area_wall_side, Q_wall_NE);
fprintf('%-15s %-12.2f %-10.2f\n','Wall NW', area_wall_front, Q_wall_NW);
fprintf('%-15s %-12.2f %-10.2f\n','Windows', area_window_total, Q_window);
fprintf('%-15s %-12.2f %-10.2f\n','Floor', area_floor, Q_floor);
fprintf('%-15s %-12.2f %-10.2f\n','Ceiling', area_ceil, Q_ceil);
fprintf('%-15s %-12s %-10.2f\n','Air Exchange', '---', Q_tot_air_exchange);

fprintf('\nTotal Heat Loss by Heat Transfer (Q_tot): %.2f W\n', Q_tot_heat_transfer);
fprintf('Total Heat Loss Including Air Exchange: %.2f W\n', Q_tot_final);



% MATLAB script to calculate heat gain (summer) including all sources: walls, windows, floor, ceiling, solar, people, lighting

% Step 1: Input shared building dimensions
fprintf('--- Building Dimensions ---\n');
b_length = input('Enter building length (m): ');
b_width = input('Enter building width (m): ');
b_height = input('Enter building height (m): ');

U_wall = input('Enter U-value for walls (W/m^2K): ');
U_floor = input('Enter U-value for floor (W/m^2K): ');
U_ceil = input('Enter U-value for ceiling (W/m^2K): ');

T_in = input('Enter internal temperature (°C): ');
T_ext = input('Enter external air temperature (°C): ');
T_ground = input('Enter ground temperature (°C): ');

deltaT_window = T_ext - T_in;
deltaT_floor = T_ground - T_in;

fprintf('\n--- deltaT_eq values for each wall and ceiling ---\n');
deltaT_eq_SW = input('Enter deltaT_eq for SW wall: ');
deltaT_eq_SE = input('Enter deltaT_eq for SE wall: ');
deltaT_eq_NE = input('Enter deltaT_eq for NE wall: ');
deltaT_eq_NW = input('Enter deltaT_eq for NW wall: ');
deltaT_eq_ceil = input('Enter deltaT_eq for ceiling: ');

% Step 2: Calculate wall areas
area_wall_side = b_width * b_height;
area_wall_front = b_length * b_height;

% Step 3: Window calculation
fprintf('\n--- Window Heat Gain Calculation ---\n');
n_windows = input('Enter number of windows: ');
win_width = input('Enter window horizontal dimension (m): ');
win_height = input('Enter window vertical dimension (m): ');
U_window = input('Enter U-value for windows (W/m^2K): ');

area_window_total = n_windows * win_width * win_height;

% Thermal conduction gain through window
Q_window = deltaT_window * area_window_total * U_window;

% Solar gain through window
I_solar = input('Enter solar irradiance on SW facade (W/m^2): ');
tau = input('Enter glazing transmittance tau (0-1): ');
Q_solar = I_solar * area_window_total * tau;

% Total window gain
Q_window_total = Q_window + Q_solar;

% Step 4: Subtract window area from SW wall and recalculate
area_wall_SW_net = area_wall_side - area_window_total;
Q_wall_SW = deltaT_eq_SW * area_wall_SW_net * U_wall;
Q_wall_SE = deltaT_eq_SE * area_wall_front * U_wall;
Q_wall_NE = deltaT_eq_NE * area_wall_side * U_wall;
Q_wall_NW = deltaT_eq_NW * area_wall_front * U_wall;

% Step 5: Floor and ceiling
area_floor = b_width * b_length;
area_ceil = area_floor;
Q_floor = deltaT_floor * area_floor * U_floor;
Q_ceil = deltaT_eq_ceil * area_ceil * U_ceil;

% Step 6: Internal heat gain from people
fprintf('\n--- Internal Heat Gain from People ---\n');
n_people = input('Enter number of people: ');
Q_person = n_people * (72.5 + 52.5); % Sensible + latent heat in W

% Step 7: Lighting gain
q_light = 12.5; % W/m^2
Q_light = q_light * area_floor;

% Step 8: Total heat gain
Q_tot_heat_transfer = Q_wall_SW + Q_wall_SE + Q_wall_NE + Q_wall_NW + Q_floor + Q_ceil + Q_window_total + Q_person + Q_light;

% Step 9: Display results
fprintf('\n%-15s %-12s %-10s\n','Element','Area(m^2)','Q_gain(W)');
fprintf('%-15s %-12.2f %-10.2f\n','Wall SW', area_wall_SW_net, Q_wall_SW);
fprintf('%-15s %-12.2f %-10.2f\n','Wall SE', area_wall_front, Q_wall_SE);
fprintf('%-15s %-12.2f %-10.2f\n','Wall NE', area_wall_side, Q_wall_NE);
fprintf('%-15s %-12.2f %-10.2f\n','Wall NW', area_wall_front, Q_wall_NW);
fprintf('%-15s %-12.2f %-10.2f\n','Windows', area_window_total, Q_window_total);
fprintf('%-15s %-12.2f %-10.2f\n','Floor', area_floor, Q_floor);
fprintf('%-15s %-12.2f %-10.2f\n','Ceiling', area_ceil, Q_ceil);
fprintf('%-15s %-12s %-10.2f\n','People', '---', Q_person);
fprintf('%-15s %-12.2f %-10.2f\n','Lighting', area_floor, Q_light);

fprintf('\nTotal Heat Gain (Q_tot): %.2f W\n', Q_tot_heat_transfer);

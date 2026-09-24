CREATE DATABASE db_demo_local;
CREATE TABLE licenses (
    id SERIAL PRIMARY KEY,
    license_key VARCHAR(50) NOT NULL UNIQUE,
    hardware_id VARCHAR(100),
    is_active BOOLEAN DEFAULT true,
    created_at TIMESTAMP WITHOUT TIME ZONE DEFAULT NOW(),
    activated_at TIMESTAMP WITHOUT TIME ZONE
);


INSERT INTO licensess2 (license_key, hardware_id, is_active, created_at, activated_at) VALUES
('HR_DEMO_KEY_20263', NULL, true, NOW(), NOW()),
('CLIENT_TEST_0023', NULL, true, NOW() - INTERVAL '1 day', NOW() - INTERVAL '1 day' + INTERVAL '5 minutes');
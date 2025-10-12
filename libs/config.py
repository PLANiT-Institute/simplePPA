"""
Configuration management utilities.
"""


def get_default_config():
    """
    Get default configuration parameters.

    Returns
    -------
    dict
        Dictionary with all configuration parameters
    """
    return {
        # Data files
        'pattern_file': 'data/pattern.xlsx',
        'kepco_file': 'data/KEPCO.xlsx',
        'kepco_year': 2024,
        'kepco_tariff': 'HV_C_III',

        # Analysis timeframe
        'start_date': '2024-01-01',
        'end_date': '2024-12-31',

        # Load capacity
        'load_capacity_mw': 3000,

        # PPA parameters
        'ppa_price': 170,          # KRW/kWh
        'ppa_mintake': 1.0,        # 1.0 = 100%, 0.5 = 50%
        'ppa_resell': False,       # Allow reselling excess energy
        'ppa_resellrate': 0.9,     # Resale rate as fraction of PPA price

        # PPA coverage range
        'ppa_range_start': 0,      # Starting percentage
        'ppa_range_end': 200,      # Ending percentage (inclusive)
        'ppa_range_step': 10,      # Step size

        # ESS parameters
        'ess_include': False,      # Include ESS in analysis
        'ess_capacity': 0.5,       # ESS capacity as percentage of solar peak (0.5 = 50%)
        'ess_price': 0.5,          # ESS discharge cost as fraction of PPA price
        'ess_hours': 6,            # ESS capacity in hours (alternative sizing method)

        # Output parameters
        'output_file': 'ppa_analysis_results.xlsx',
        'verbose': False,          # Print detailed statistics during calculation
        'export_long_format': True,  # Export long-format data for pivot tables
    }


def load_config_from_file(filepath):
    """
    Load configuration from file.

    Parameters
    ----------
    filepath : str
        Path to configuration file (JSON or YAML)

    Returns
    -------
    dict
        Configuration dictionary
    """
    import json
    import os

    config = get_default_config()

    if not os.path.exists(filepath):
        return config

    # Load from JSON
    if filepath.endswith('.json'):
        with open(filepath, 'r') as f:
            user_config = json.load(f)
            config.update(user_config)

    # Load from YAML (if pyyaml is available)
    elif filepath.endswith('.yaml') or filepath.endswith('.yml'):
        try:
            import yaml
            with open(filepath, 'r') as f:
                user_config = yaml.safe_load(f)
                config.update(user_config)
        except ImportError:
            print("Warning: pyyaml not installed, cannot load YAML config")

    return config


def save_config_to_file(config, filepath):
    """
    Save configuration to file.

    Parameters
    ----------
    config : dict
        Configuration dictionary
    filepath : str
        Path to save configuration file (JSON or YAML)

    Returns
    -------
    None
    """
    import json

    # Save as JSON
    if filepath.endswith('.json'):
        with open(filepath, 'w') as f:
            json.dump(config, f, indent=4)

    # Save as YAML (if pyyaml is available)
    elif filepath.endswith('.yaml') or filepath.endswith('.yml'):
        try:
            import yaml
            with open(filepath, 'w') as f:
                yaml.dump(config, f, default_flow_style=False)
        except ImportError:
            print("Warning: pyyaml not installed, cannot save YAML config")


def validate_config(config):
    """
    Validate configuration parameters.

    Parameters
    ----------
    config : dict
        Configuration dictionary

    Returns
    -------
    tuple
        (is_valid, error_messages)
    """
    errors = []

    # Validate ranges
    if config['ppa_range_start'] < 0:
        errors.append("ppa_range_start must be >= 0")

    if config['ppa_range_end'] < config['ppa_range_start']:
        errors.append("ppa_range_end must be >= ppa_range_start")

    if config['ppa_range_step'] <= 0:
        errors.append("ppa_range_step must be > 0")

    # Validate PPA parameters
    if config['ppa_price'] <= 0:
        errors.append("ppa_price must be > 0")

    if not (0 <= config['ppa_mintake'] <= 1):
        errors.append("ppa_mintake must be between 0 and 1")

    if not (0 <= config['ppa_resellrate'] <= 1):
        errors.append("ppa_resellrate must be between 0 and 1")

    # Validate ESS parameters
    if config['ess_capacity'] < 0:
        errors.append("ess_capacity must be >= 0")

    if config['ess_price'] < 0:
        errors.append("ess_price must be >= 0")

    # Validate load capacity
    if config['load_capacity_mw'] <= 0:
        errors.append("load_capacity_mw must be > 0")

    # Validate dates
    try:
        from datetime import datetime
        datetime.strptime(config['start_date'], '%Y-%m-%d')
        datetime.strptime(config['end_date'], '%Y-%m-%d')
    except ValueError:
        errors.append("Dates must be in YYYY-MM-DD format")

    return len(errors) == 0, errors


def print_config(config):
    """
    Print configuration in readable format.

    Parameters
    ----------
    config : dict
        Configuration dictionary

    Returns
    -------
    None
    """
    print("\n=== CONFIGURATION ===")
    print("\nData Files:")
    print(f"  Pattern file: {config['pattern_file']}")
    print(f"  KEPCO file: {config['kepco_file']}")
    print(f"  KEPCO tariff: {config['kepco_tariff']} ({config['kepco_year']})")

    print("\nAnalysis Period:")
    print(f"  Start: {config['start_date']}")
    print(f"  End: {config['end_date']}")

    print("\nLoad Parameters:")
    print(f"  Capacity: {config['load_capacity_mw']} MW")

    print("\nPPA Parameters:")
    print(f"  Price: {config['ppa_price']} KRW/kWh")
    print(f"  Minimum take: {config['ppa_mintake']*100:.0f}%")
    print(f"  Resell allowed: {config['ppa_resell']}")
    if config['ppa_resell']:
        print(f"  Resell rate: {config['ppa_resellrate']*100:.0f}%")

    print("\nPPA Coverage Range:")
    print(f"  Start: {config['ppa_range_start']}%")
    print(f"  End: {config['ppa_range_end']}%")
    print(f"  Step: {config['ppa_range_step']}%")
    print(f"  Scenarios: {len(range(config['ppa_range_start'], config['ppa_range_end']+1, config['ppa_range_step']))}")

    print("\nESS Parameters:")
    print(f"  Include ESS: {config['ess_include']}")
    if config['ess_include']:
        print(f"  Capacity: {config['ess_capacity']*100:.0f}% of peak solar")
        print(f"  Discharge price: {config['ess_price']*100:.0f}% of PPA price")

    print("\nOutput:")
    print(f"  File: {config['output_file']}")
    print(f"  Long format: {config['export_long_format']}")
    print(f"  Verbose: {config['verbose']}")

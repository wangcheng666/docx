
import xml.etree.ElementTree as ET

def is_namespace_registered(prefix: str) -> bool:
    """
    Check if the provided namespace is registered.
    """
    # Check if the provided namespace is in the list of registered namespaces
    return prefix in ET._namespace_map.keys()

def register_namespace(prefix: str, uri: str):
    """
    Register a new namespace.
    """
    ET.register_namespace(prefix, uri)

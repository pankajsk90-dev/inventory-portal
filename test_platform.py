import urllib.request
import json

def test():
    for plat in ['Blinkit', 'Zepto', 'Amazon', 'Instamart']:
        url = f'http://127.0.0.1:8765/api/platform-detail?platform={plat}'
        with urllib.request.urlopen(url) as response:
            data = json.loads(response.read().decode('utf-8'))
        print(f"=== Platform: {data['platform']} ===")
        print(f"  Total Units Sent: {data['total_units_sent']:,}")
        print(f"  Total Dispatches: {data['total_dispatches']}")
        print(f"  PO Fill Rate: {data['fill_rate']}%")
        top_prod = data['products'][0] if data['products'] else None
        if top_prod:
            print(f"  Top Product: {top_prod['product_name']} ({top_prod['units_sent']:,} units, {top_prod['share_percent']}%)")
        print()

if __name__ == '__main__':
    test()

from utils.descritivePivotUtils import descritivePivot


if __name__ == "__main__":
    desc_pivot_obj = descritivePivot("config/descriptive_pivot_conf.json")
    desc_pivot_obj.create_descritive_pivot()

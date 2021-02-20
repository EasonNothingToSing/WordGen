import WordGen
import argparse


if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog="WordGen",
                                     description="介绍： ListenAI 旗下 chipsky 的专用文档生成器，"
                                                 "根据IC生成的XLS文件，生成模板，再由模板渲染成最终文档."
                                                 "\r\n局限性： XLS文件中关键字发生改变的时候，"
                                                 "会导致模板发生变化，从而无法对原有模板进行渲染."
                                                 "\r\n作者： EASON WANG",
                                     epilog="eg: --update --generate,\r\n 尽量不要使用 include和exclude参数",
                                     formatter_class=argparse.RawTextHelpFormatter)

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--generate", "-g", action="store_true")
    group.add_argument("--update", "-u", action="store_true")
    group.add_argument("--verify", "-v", action="store_true")

    parser.add_argument("--include", "-i", nargs="+", type=str)
    parser.add_argument("--exclude", "-e", nargs="+", type=str)
    parser.add_argument("--version", action="version", version="%%(prog)s %s" % WordGen.WORDGEN_VERSION)

    args = parser.parse_args("--update".split())

    if args.generate:
        # 生成word模板
        WordGen.generate_template(args.include, args.exclude)
    elif args.update:
        # 渲染word模板
        WordGen.update_template(args.include, args.exclude)
    elif args.verify:
        # 检查excel文档是否命名正确
        WordGen.excel_content_verify(args.include, args.exclude)

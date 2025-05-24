import CCX_Sheet


def main():
    banner = """ 
*********************************************************                                 
         __  __       _____             __  __  
    |  |/  `/  `\_/    |/  \   |  |\_/ /  `/  ` 
    \__/\__,\__,/ \    |\__/   |/\|/ \ \__,\__,                                                                 

*********************************************************
*                                                       *
*   Welcome to the UCCX to WxCC Migration tool!         *
*   Created by Simranjit Singh                          *
*                                                       *
*   For more updates, visit:                            *
*   https://learnuccollab.com/                          *
*                                                       *
*********************************************************
"""
    print(f"{banner}")
    print("This tool (works with Windows/Linux) is designed to help you migrate the following Voice channel configurations and data from Cisco UCCX to Cisco WxCC in one go.")
    print("1. Applications\n"
          "2. Teams\n"
          "3. Trigger\n"
          "4. Skills\n"
          "5. CSQs\n"
          "6. Wrap-Up Codes\n"
          "7. Reason Codes\n"
          "8. Phonebooks")

    ques_1 = input("\033[94mAre you ready to continue (Y/N) ? : \033[0m")
    if ques_1 == "Y" or ques_1 == "y":
        print("\033[93mCapturing configuration details from UCCX.....\033[0m")
        CCX_Sheet.get_APP()
    else:
        print("Closing the app.")

if __name__ == '__main__':
    main()
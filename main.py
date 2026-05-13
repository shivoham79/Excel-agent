"""Entry point — interactive chat loop with the Excel agent."""

from dotenv import load_dotenv

load_dotenv()

from agent import create_agent


def main():
    agent = create_agent()
    print("Excel Agent ready. Type 'quit' to exit.\n")

    while True:
        user_input = input("You: ").strip()
        if user_input.lower() in ("quit", "exit", "q"):
            print("Bye!")
            break
        if not user_input:
            continue

        response = agent.invoke({"messages": [("human", user_input)]})
        # Last message from the agent
        ai_msg = response["messages"][-1]
        print(f"\nAgent: {ai_msg.content}\n")


if __name__ == "__main__":
    main()

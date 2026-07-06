name = input("Apna naam likho: ")
age = int(input("Apni age likho: "))
password = input("Password enter karo: ")

print("\nChecking... 🔍")

if password == "python123":
    print("✅ Password sahi hai!")

    if age >= 18:
        print(f"🎉 Welcome {name}! Tum voting bhi kar sakte ho.")
    else:
        print(f"😄 Welcome {name}! Abhi thoda aur bade ho jao.")

else:
    print("❌ Password galat hai!")
    print("🚔 FBI: Hum tumhari location track kar rahe hain... 😂")